"""Verify release checksums, source allowlist inventory and portable payload hashes."""
import argparse
import hashlib
import json
import sys
import zipfile
from pathlib import Path, PurePosixPath

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from release_integrity import payload_name


def digest(stream):
    result = hashlib.sha256()
    for chunk in iter(lambda: stream.read(1024 * 1024), b''):
        result.update(chunk)
    return result.hexdigest()


def check_archive(archive, expected):
    entries = archive.infolist()
    assert len(entries) == len({entry.filename for entry in entries}), 'Duplicate ZIP entries'
    assert {entry.filename for entry in entries} == set(expected), 'Archive inventory mismatch'
    for entry in entries:
        payload_name(entry.filename)
        assert (entry.external_attr >> 16) & 0o170000 != 0o120000, 'Linked ZIP entry'
        record = expected[entry.filename]
        assert entry.file_size == record['size'], entry.filename
        with archive.open(entry) as stream:
            assert digest(stream) == record['sha256'], entry.filename


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('release_dir', type=Path)
    release = parser.parse_args().release_dir.resolve()
    manifest = json.loads((release / 'RELEASE_MANIFEST.json').read_text(encoding='utf-8'))
    for record in manifest['files']:
        payload_name(record['path'])
        path = (release / record['path']).resolve()
        assert path.is_relative_to(release), 'Manifest path escapes release'
        assert path.stat().st_size == record['size'], path
        with path.open('rb') as stream:
            assert digest(stream) == record['sha256'], path
    source_archive = release / 'FINDEX_v3.0.1_Source.zip'
    if source_archive.is_file():
        with zipfile.ZipFile(source_archive) as archive:
            source = json.loads(archive.read('SOURCE_FILES.json'))
            expected = {record['path']: record for record in source}
            assert len(expected) == len(source), 'Duplicate source records'
            for metadata in ('SOURCE_FILES.json', 'NATIVE_COMPONENTS.json'):
                data = archive.read(metadata)
                expected[metadata] = dict(size=len(data), sha256=hashlib.sha256(data).hexdigest())
            for name in expected:
                path = PurePosixPath(name)
                assert path.suffix.lower() not in ('.exe', '.dll', '.pyd', '.bak', '.pem', '.key', '.pfx', '.p12'), name
                assert not any(part in ('.venv', 'scratch', 'Metadata_Reports', '__pycache__', '.git', '.aws') for part in path.parts), name
            check_archive(archive, expected)
            print('Verified source inventory and exclusion rules:', len(expected), 'files')
    else:
        from prepare_github import validate
        folder = Path(__file__).resolve().parents[1] / 'github/FINDEX'
        expected = json.loads((folder / 'SOURCE_FILES.json').read_text(encoding='utf-8'))
        assert validate(folder) == expected, 'GitHub source inventory mismatch'
        print('Verified GitHub source folder:', len(expected) + 1, 'files')
    root = Path(__file__).resolve().parents[1]
    payload_path = release / 'APPLICATION_PAYLOAD.json'
    if not payload_path.is_file():
        payload_path = root / 'dist/app_payload.manifest.json'
    payload = json.loads(payload_path.read_text(encoding='utf-8'))
    portable = release / 'FINDEX_v3.0.1_Portable.zip'
    if portable.is_file():
        with zipfile.ZipFile(portable) as archive:
            check_archive(archive, {'FINDEX/' + name: record for name, record in payload['files'].items()})
        print('Verified portable ZIP against every payload hash:', len(payload['files']), 'files')
    else:
        application = root / 'dist/FINDEX'
        actual = {p.relative_to(application).as_posix() for p in application.rglob('*') if p.is_file()}
        assert actual == set(payload['files']), 'Application folder inventory mismatch'
        for name, record in payload['files'].items():
            path = (application / name).resolve()
            assert path.is_relative_to(application.resolve()), 'Unsafe application path'
            assert path.stat().st_size == record['size'], name
            with path.open('rb') as stream:
                assert digest(stream) == record['sha256'], name
        print('Verified retained application folder:', len(actual), 'files')
    print('Verified all', len(manifest['files']), 'release checksums; publisher authenticity is not implied.')


if __name__ == '__main__':
    main()
