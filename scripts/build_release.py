"""Build and verify FINDEX without changing Windows script/security policies."""
import argparse
import json
import os
import subprocess
import sys
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--environment', default='.venv')
parser.add_argument('--output-dir', default='release/FINDEX-v3.0.1', help='New, empty release directory inside this checkout')
arguments = parser.parse_args()
release = (ROOT / arguments.output_dir).resolve()
if not release.is_relative_to(ROOT) or release == ROOT:
    raise SystemExit('Release directory must be inside the checkout')
if release.exists() and any(release.iterdir()):
    raise SystemExit('Use a new or empty release directory; existing releases are preserved')
release.mkdir(parents=True, exist_ok=True)
python = ROOT / arguments.environment / 'Scripts/python.exe'
if not python.is_file():
    raise SystemExit('Create an isolated Python 3.12 environment first')
config = python.parent.parent / 'pyvenv.cfg'
if 'include-system-site-packages = false' not in config.read_text():
    raise SystemExit('The build environment must not inherit system packages')
environment = os.environ.copy()
environment.pop('PYTHONPATH', None)
environment['PYTHONNOUSERSITE'] = '1'
if os.name == 'nt':
    windows = Path(os.environ.get('SystemRoot', r'C:\Windows'))
    environment['PATH'] = os.pathsep.join((str(windows / 'System32'), str(windows), environment.get('PATH', '')))


def run(*args):
    subprocess.run([str(python), *args], cwd=ROOT, env=environment, check=True)


run('scripts/check_build_environment.py')
run('scripts/collect_notices.py')
run('tests/test_security_hardening.py')
run('tests/test_gui_corrections.py')
run('tests/test_message_box_theme.py')
run('tests/test_report_branding.py')
run('tests/test_export_selection.py')
run('tests/test_inventory_sizes.py')
run('tests/check_startup_windows.py')
run('-m', 'PyInstaller', '--clean', '--noconfirm', 'File_Inventory_Pro.spec')
run('scripts/check_binary_identity.py', 'dist/FINDEX/FINDEX.exe')
run('tests/smoke_portable.py')
run('scripts/package_payload.py')
run('tests/test_payload_install.py')
run('-m', 'PyInstaller', '--clean', '--noconfirm', '--distpath', 'scratch/setup-stage', 'installer_build.spec')
run('scripts/check_binary_identity.py', 'scratch/setup-stage/FINDEX_v3.0.1_Setup.exe')
run('tests/smoke_installer.py', str(ROOT / 'scratch/setup-stage/FINDEX_v3.0.1_Setup.exe'))
shutil.copy2(ROOT / 'scratch/setup-stage/FINDEX_v3.0.1_Setup.exe', release / 'FINDEX_v3.0.1_Setup.exe')
run('scripts/prepare_public_release.py', '--output-dir', str(release))
run('scripts/release_manifest.py', '--release-dir', str(release))
sys.path.insert(0, str(ROOT / 'scripts'))
from release_manifest import sha256
manifest = json.loads((release / 'RELEASE_MANIFEST.json').read_text(encoding='utf-8'))
for entry in manifest['files']:
    path = (release / entry['path']).resolve()
    if not path.is_relative_to(release) or sha256(path) != entry['sha256']:
        raise SystemExit('Release verification failed: ' + entry['path'])
print('Verified all', len(manifest['files']), 'release checksums. Complete the remaining public release gates in SECURITY.md.')
run('scripts/validate_public_release.py', str(release))
