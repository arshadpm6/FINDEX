"""Inventory release artifacts, tool versions and installed build dependencies."""
import datetime
import argparse
import hashlib
import importlib.metadata
import json
import platform
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def sha256(path):
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--release-dir', type=Path, default=ROOT)
    release = parser.parse_args().release_dir.resolve()
    if not release.is_relative_to(ROOT):
        raise SystemExit('Release directory must be inside this checkout')
    application = ROOT / 'dist' / 'FINDEX'
    setup = release / 'FINDEX_v3.0.1_Setup.exe'
    if not setup.is_file() or not (application / 'FINDEX.exe').is_file():
        raise SystemExit('Build both FINDEX and Setup first')
    internal = application / '_internal'
    perl = internal / 'external/exiftool/exiftool_files'
    exif = subprocess.run([str(perl / 'perl.exe'), '-I', str(perl / 'lib'),
        str(perl / 'exiftool.pl'), '-config', '', '-ver'], check=True, capture_output=True,
        text=True, timeout=20, creationflags=subprocess.CREATE_NO_WINDOW).stdout.strip()
    probe = subprocess.run([str(internal / 'external/ffmpeg/ffprobe.exe'), '-version'],
        check=True, capture_output=True, text=True, timeout=20,
        creationflags=subprocess.CREATE_NO_WINDOW).stdout
    paths = sorted(p for p in release.rglob('*') if p.is_file()
                   and p.name not in ('RELEASE_MANIFEST.json', 'SHA256SUMS.txt')) if release != ROOT else [setup, *sorted(p for p in application.rglob('*') if p.is_file())]
    files = [dict(path=p.relative_to(release).as_posix(), size=p.stat().st_size, sha256=sha256(p)) for p in paths]
    document = dict(application='FINDEX', version='3.0.1',
        built_at_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        python=platform.python_version(), platform=platform.platform(),
        build_dependencies=sorted([dict(name=d.metadata['Name'], version=d.version)
            for d in importlib.metadata.distributions()], key=lambda d: d['name'].lower()),
        exiftool_version=exif, ffprobe_version_and_configuration=probe,
        signing='Unsigned; no publisher authentication is provided by these checksums.',
        files=files)
    (release / 'RELEASE_MANIFEST.json').write_text(json.dumps(document, indent=2), encoding='utf-8')
    (release / 'SHA256SUMS.txt').write_text(''.join(f"{item['sha256']}  {item['path']}\n" for item in files), encoding='utf-8')
    print(f'Recorded {len(files)} release files; ExifTool {exif}')


if __name__ == '__main__':
    main()
