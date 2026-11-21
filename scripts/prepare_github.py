"""Create a new source-only GitHub folder; never upload, commit or include native binaries."""
import argparse
import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parents[1]


def sha(data):
    return hashlib.sha256(data).hexdigest()


def validate(folder):
    forbidden = {'.exe', '.dll', '.pyd', '.so', '.dylib', '.lib', '.a', '.whl', '.zip',
                 '.7z', '.bak', '.pem', '.key', '.pfx', '.p12', '.xlsx', '.csv'}
    records = []
    for path in sorted(folder.rglob('*')):
        if '.git' in path.relative_to(folder).parts:
            continue
        if path.is_symlink() or path.is_junction():
            raise ValueError('Linked publication path: ' + str(path))
        if not path.is_file():
            continue
        name = path.relative_to(folder).as_posix()
        assert not any(part in {'.venv', 'scratch', 'Metadata_Reports', '__pycache__', '.aws', '.codex'}
                       for part in path.relative_to(folder).parts), name
        assert path.suffix.lower() not in forbidden, name
        data = path.read_bytes()
        assert not data.startswith((b'MZ', b'\x7fELF')), 'Native binary: ' + name
        if path.suffix == '.py':
            compile(data, name, 'exec')
        if path.suffix in {'.py', '.md', '.json', '.ps1', '.yml', '.txt'}:
            text = data.decode('utf-8-sig')
            assert not re.search(r'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----|\bgh[pousr]_[A-Za-z0-9]{30,}|\bAKIA[A-Z0-9]{16}\b', text), 'Possible credential: ' + name
        if name != 'SOURCE_FILES.json':
            records.append(dict(path=name, size=len(data), sha256=sha(data)))
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'github/FINDEX')
    parser.add_argument('--release-dir', type=Path, default=ROOT / 'release/FINDEX-v3.0.1')
    args = parser.parse_args()
    folder = args.output_dir.resolve()
    assert folder.is_relative_to(ROOT) and folder != ROOT, 'Output must stay inside checkout'
    assert not folder.exists() or not any(folder.iterdir()), 'Use a new or empty GitHub folder'
    archive_path = args.release_dir.resolve() / 'FINDEX_v3.0.1_Source.zip'
    assert archive_path.is_relative_to(ROOT), 'Release input must stay inside checkout'
    folder.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive_path) as archive:
        inventory = {r['path']: r for r in json.loads(archive.read('SOURCE_FILES.json'))}
        expected = set(inventory) | {'SOURCE_FILES.json', 'NATIVE_COMPONENTS.json'}
        names = archive.namelist()
        assert set(names) == expected and len(names) == len(expected), 'Unexpected archive files'
        for name in names:
            relative = PurePosixPath(name)
            assert not relative.is_absolute() and '..' not in relative.parts and ':' not in name and '\\' not in name
            data = archive.read(name)
            if name in inventory:
                assert sha(data) == inventory[name]['sha256'], name
            if name == 'SOURCE_FILES.json':
                continue  # Recomputed after the source-only documentation overlays.
            target = folder.joinpath(*relative.parts)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
    # Reviewed publishing material is separate from the executable release snapshot.
    for source, target in [('publishing/GITHUB_README.md', 'README.md'),
                           ('publishing/SOURCE_INSTALL.md', 'INSTALL.md'),
                           ('publishing/SOURCE_RELEASE_NOTES.md', 'RELEASE_NOTES.md'),
                           ('publishing/GITHUB_PUBLISHING.md', 'PUBLIC_RELEASE.md'),
                           ('publishing/SOURCE_SECURITY.md', 'SECURITY.md'),
                           ('NATIVE_DISTRIBUTION_REVIEW.md', 'NATIVE_DISTRIBUTION_REVIEW.md'),
                           ('NATIVE_DISTRIBUTION_REVIEW.json', 'NATIVE_DISTRIBUTION_REVIEW.json'),
                           ('publishing/CONTRIBUTING.md', 'CONTRIBUTING.md'),
                           ('scripts/prepare_github.py', 'scripts/prepare_github.py')]:
        shutil.copy2(ROOT / source, folder / target)
    shutil.copytree(ROOT / 'publishing', folder / 'publishing', dirs_exist_ok=True)
    with (folder / '.gitignore').open('a', encoding='utf-8') as stream:
        stream.write('\n# Native tools, report exports and public-source staging must not be committed.\n'
                     'github/\n*.dll\n*.pyd\n*.lib\n*.a\n*.so\n*.dylib\n*.whl\n'
                     '*.xlsx\n*.xls\n*.csv\n*.tsv\n*.html\n')
    records = validate(folder)
    (folder / 'SOURCE_FILES.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
    print('Validated source-only GitHub folder:', folder, 'files:', len(records) + 1)
    print('No Git commit, push, LinkedIn post or binary publication was performed.')


if __name__ == '__main__':
    main()
