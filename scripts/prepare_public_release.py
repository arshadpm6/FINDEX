"""Prepare allowlisted public-source and portable archives; never publish them."""
import argparse
import hashlib
import json
import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ('README.md', 'LICENSE', 'INSTALL.md', 'BUILD.md', 'CHANGELOG.md',
        'SECURITY.md', 'THIRD_PARTY_NOTICES.md', 'RELEASE_NOTES.md',
        'PUBLIC_RELEASE.md', 'FINAL_RELEASE_VERIFICATION.md')
SOURCES = ('File Inventory Pro.py', 'installer_wizard.py', 'security_io.py',
           'release_integrity.py', 'dialog_theme.py', 'report_branding.py',
           'excel_support.py', 'excel_diagnostics.py',
           'File_Inventory_Pro.spec', 'installer_build.spec', '.gitignore',
           'requirements.txt', 'requirements-build.txt', 'requirements-release.lock',
           'logo.png', 'icon.png', 'process.png', 'app_icon.ico')


def digest(path):
    value = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            value.update(block)
    return value.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', required=True, type=Path)
    parser.add_argument('--source-only', action='store_true', help='Refresh source/docs while preserving an already verified portable ZIP')
    args = parser.parse_args()
    release = args.output_dir.resolve()
    if not release.is_relative_to(ROOT) or release == ROOT:
        raise SystemExit('Output must be inside the checkout')
    release.mkdir(parents=True, exist_ok=True)
    selected = [ROOT / name for name in (*DOCS, *SOURCES)]
    for folder, extensions in (('scripts', {'.py', '.ps1'}), ('tests', {'.py'}),
                               ('assets', {'.png', '.svg', '.ico'}),
                               ('resources/branding', {'.png', '.ico', '.txt', '.md'})):
        selected += [p for p in (ROOT / folder).rglob('*')
                     if p.is_file() and p.suffix.lower() in extensions and '__pycache__' not in p.parts]
    selected += [ROOT / 'external/ffmpeg' / name for name in ('LICENSE', 'README.txt', 'SOURCE.json')]
    selected += [p for p in (ROOT / 'resources/third_party/Qt-license-texts').rglob('*') if p.is_file()]
    selected = sorted(set(selected))
    native = [dict(path=p.relative_to(ROOT).as_posix(), size=p.stat().st_size, sha256=digest(p))
              for p in sorted((ROOT / 'external').rglob('*')) if p.is_file()]
    native_json = json.dumps(dict(note='Native tools excluded from source ZIP; provision separately. These are local hashes, not upstream signatures.', files=native), indent=2)
    source_files = [dict(path=p.relative_to(ROOT).as_posix(), size=p.stat().st_size, sha256=digest(p)) for p in selected]
    with zipfile.ZipFile(release / 'FINDEX_v3.0.1_Source.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
        for path in selected:
            archive.write(path, path.relative_to(ROOT).as_posix())
        archive.writestr('NATIVE_COMPONENTS.json', native_json)
        archive.writestr('SOURCE_FILES.json', json.dumps(source_files, indent=2))
    for name in DOCS:
        shutil.copy2(ROOT / name, release / name)
    audit = ROOT / 'scratch/final_release_dependency_audit.json'
    if audit.is_file():
        shutil.copy2(audit, release / 'DEPENDENCY_AUDIT.json')
    payload = ROOT / 'dist/app_payload.manifest.json'
    if payload.is_file():
        shutil.copy2(payload, release / 'APPLICATION_PAYLOAD.json')
    if args.source_only:
        if not (release / 'FINDEX_v3.0.1_Portable.zip').is_file():
            raise SystemExit('No existing portable ZIP to preserve')
        print('Refreshed source archive and release documentation:', release)
        return
    with zipfile.ZipFile(release / 'FINDEX_v3.0.1_Portable.zip', 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted((ROOT / 'dist/FINDEX').rglob('*')):
            if path.is_file():
                archive.write(path, path.relative_to(ROOT / 'dist').as_posix())
    print('Prepared local portable/source archives and release documentation:', release)


if __name__ == '__main__':
    main()
