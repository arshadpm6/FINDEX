"""Check the rebuilt portable program and its bundled helper executables."""
import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
import tempfile
import shutil
from pathlib import Path

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('portable', nargs='?', type=Path, default=root / 'dist/FINDEX')
parser.add_argument('--core', action='store_true')
options = parser.parse_args()
portable = options.portable.resolve()
executable = portable / 'FINDEX.exe'
internal = portable / '_internal'
tools = [internal / 'external' / 'exiftool' / 'exiftool.exe', internal / 'external' / 'ffmpeg' / 'ffprobe.exe']
assert executable.is_file()
with tempfile.TemporaryDirectory(prefix='findex-frozen-excel-', dir=root / 'scratch') as temporary:
    temporary_root = Path(temporary).resolve()
    assert temporary_root.is_relative_to((root / 'scratch').resolve())
    for name in ('clean', 'legacy-numpy'):
        stub = internal / 'numpy'
        if name == 'legacy-numpy':
            assert not stub.exists(), 'Do not overwrite a real NumPy installation'
            stub.mkdir()
            (stub / '__init__.py').write_text('# FINDEX temporary incomplete optional NumPy fixture\n', encoding='utf-8')
        try:
            output = temporary_root / name
            checked = subprocess.run([str(executable), '--check-excel', str(output)],
                cwd=portable, capture_output=True, timeout=90, creationflags=subprocess.CREATE_NO_WINDOW)
            result_path = output / 'result.json'
            assert checked.returncode == 0, (checked.returncode,
                result_path.read_text(encoding='utf-8') if result_path.exists() else checked.stderr)
            excel_result = json.loads(result_path.read_text(encoding='utf-8'))
            assert excel_result['status'] == 'passed' and excel_result['excel_available']
            assert len(excel_result['reports']) == 7
            if name == 'legacy-numpy':
                assert excel_result['numpy_fallback'] and 'numpy' in excel_result['numpy_fallback']
            print('PASS: actual frozen Excel/HTML/hash exports:', name, flush=True)
        finally:
            if name == 'legacy-numpy':
                assert stub.resolve().is_relative_to(internal.resolve()) and stub.name == 'numpy'
                assert (stub / '__init__.py').read_text(encoding='utf-8').startswith('# FINDEX temporary')
                shutil.rmtree(stub)  # Only the explicitly owned fixture in the freshly built app.
versions = {}
if options.core:
    assert not (internal / 'external').exists(), 'Core package must not include helper bundles'
    tools = []
for tool in tools:
    assert tool.is_file(), tool
    if tool.name == 'exiftool.exe':
        runtime = tool.parent / 'exiftool_files'
        command = [str(runtime / 'perl.exe'), '-I', 'lib', 'exiftool.pl', '-config', '', '-ver']
        working = runtime
    else:
        command, working = [str(tool), '-version'], portable
    result = subprocess.run(command, cwd=working, shell=False,
                            capture_output=True, text=True, timeout=20, creationflags=subprocess.CREATE_NO_WINDOW)
    assert result.returncode == 0, result.stderr
    versions[tool.name] = result.stdout.splitlines()[0]
environment = os.environ.copy()
environment['QT_QPA_PLATFORM'] = 'offscreen'
process = subprocess.Popen([str(executable)], cwd=portable, env=environment, creationflags=subprocess.CREATE_NO_WINDOW)
try:
    deadline = time.monotonic() + 6
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f'Portable application exited during startup: {process.returncode}')
        time.sleep(.1)
finally:
    if process.poll() is None:
        process.terminate()
        process.wait(10)
report = dict(portable_executable=str(executable), startup='passed (offscreen)', helper_versions=versions,
              executable_sha256=hashlib.sha256(executable.read_bytes()).hexdigest(),
              source_sha256=hashlib.sha256((root / 'File Inventory Pro.py').read_bytes()).hexdigest())
report_dir = root / 'scratch' / 'verified_inventory_output'
report_dir.mkdir(parents=True, exist_ok=True)
(report_dir / 'portable_verification.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
