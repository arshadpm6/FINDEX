"""Build the installer payload from the freshly built portable application."""
import argparse
from pathlib import Path
import zipfile
import shutil
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from release_integrity import payload_manifest
import json

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument(
    '--application-dir',
    type=Path,
    default=root / 'dist' / 'FINDEX',
    help='One-folder PyInstaller application directory',
)
args = parser.parse_args()
application = args.application_dir.resolve()
required_files = (
    application / 'FINDEX.exe',
    application / '_internal' / 'PySide6' / 'QtWidgets.pyd',
)
missing = [str(path) for path in required_files if not path.is_file()]
if missing:
    raise SystemExit('Incomplete application build; missing: ' + ', '.join(missing))
output = root / 'dist' / 'app_payload.zip'
for name in ('README.md', 'LICENSE', 'SECURITY.md', 'THIRD_PARTY_NOTICES.md',
             'CHANGELOG.md', 'INSTALL.md', 'BUILD.md', 'RELEASE_NOTES.md',
             'PUBLIC_RELEASE.md', 'FINAL_RELEASE_VERIFICATION.md'):
    shutil.copy2(root / name, application / name)
with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
    for source in sorted(application.rglob('*')):
        if source.is_file():
            archive.write(source, source.relative_to(application).as_posix())
print(output, output.stat().st_size, 'bytes')
with zipfile.ZipFile(output) as archive:
    manifest = payload_manifest(archive)
output.with_suffix('.manifest.json').write_text(json.dumps(manifest, indent=2), encoding='utf-8')
