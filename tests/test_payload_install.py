"""Exercise the real payload in a disposable Unicode installation directory."""
import argparse
import hashlib
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import time
import wave
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument('--core', action='store_true')
options = parser.parse_args()
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ['QT_QPA_PLATFORM'] = 'offscreen'
from PySide6.QtWidgets import QApplication

spec = importlib.util.spec_from_file_location('payload_installer', ROOT / 'installer_wizard.py')
installer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(installer)
app = QApplication.instance() or QApplication([])
(ROOT / 'scratch').mkdir(exist_ok=True)
with tempfile.TemporaryDirectory(prefix='findex-installed-é-', dir=ROOT / 'scratch') as directory:
    temporary = Path(directory).resolve()
    assert temporary.is_relative_to(ROOT / 'scratch')
    destination = temporary / 'FINDEX'
    results = []
    worker = installer.ExtractWorker(ROOT / 'dist/app_payload.zip', destination, False, False, False)
    worker.finished.connect(lambda *args: results.append(args))
    worker.run()
    assert results and results[0][0], results
    print('PASS: actual payload integrity and extraction to Unicode installation path', flush=True)
    report = destination / 'user report é.txt'
    report.write_text('Keep this user report', encoding='utf-8')
    subprocess.run([sys.executable, str(ROOT / 'tests/smoke_portable.py'), str(destination)] + (['--core'] if options.core else []),
                   cwd=ROOT, check=True, timeout=120)
    source = temporary / 'नमूना é.wav'
    with wave.open(str(source), 'wb') as output:
        output.setnchannels(1)
        output.setsampwidth(2)
        output.setframerate(8000)
        output.writeframes(b'\x00\x00' * 8000)
    before = hashlib.sha256(source.read_bytes()).hexdigest(), source.stat().st_mtime_ns
    internal = destination / '_internal'
    if options.core:
        assert not (internal / 'external').exists()
        print('PASS: installed core package contains no optional helper bundles', flush=True)
    else:
        runtime = internal / 'external/exiftool/exiftool_files'
        output = subprocess.run([str(runtime / 'perl.exe'), '-I', 'lib', 'exiftool.pl', '-config', '',
            '-charset', 'filename=UTF8', '-json', '-@', '-'], input=str(source) + '\n',
            cwd=runtime, capture_output=True, text=True, encoding='utf-8', check=True,
            timeout=45, creationflags=subprocess.CREATE_NO_WINDOW)
        assert json.loads(output.stdout)[0]['FileType'] == 'WAV'
        duration = subprocess.run([str(internal / 'external/ffmpeg/ffprobe.exe'), '-v', 'error',
            '-protocol_whitelist', 'file,pipe', '-show_entries', 'format=duration', '-of',
            'default=noprint_wrappers=1:nokey=1', str(source)], capture_output=True, text=True,
            check=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW)
        assert abs(float(duration.stdout.strip()) - 1) < .01
        assert (hashlib.sha256(source.read_bytes()).hexdigest(), source.stat().st_mtime_ns) == before
        print('PASS: installed ExifTool/ffprobe process Unicode evidence without changing content or mtime', flush=True)
    # Only this fixture's own manifest/files are eligible for deletion. No shortcuts or registry were created.
    uninstall_started = time.monotonic()
    uninstall = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy',
        'Bypass', '-File', str(destination / 'uninstall.ps1')], capture_output=True, text=True, timeout=600)
    assert uninstall.returncode == 0, uninstall.stderr
    assert report.read_text(encoding='utf-8') == 'Keep this user report'
    assert not (destination / 'FINDEX.exe').exists()
    assert source.is_file()
    print(f'PASS: real-payload uninstall retains user report and evidence fixture ({time.monotonic() - uninstall_started:.2f}s)', flush=True)
print('Local payload checks passed; this is not a clean Windows VM test.', flush=True)
