"""Check exact Setup startup offscreen without pressing Install or changing settings."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]
executable = Path(sys.argv[1]).resolve()
assert executable.is_file() and executable.is_relative_to(ROOT)
temporary = tempfile.TemporaryDirectory(prefix='findex-setup-smoke-', dir=ROOT / 'scratch')
directory = Path(temporary.name).resolve()
assert directory.is_relative_to((ROOT / 'scratch').resolve())
environment = os.environ.copy()
environment.update(QT_QPA_PLATFORM='offscreen', TEMP=str(directory), TMP=str(directory))
process = subprocess.Popen([str(executable)], cwd=executable.parent, env=environment,
                           creationflags=subprocess.CREATE_NO_WINDOW,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
try:
    deadline = time.monotonic() + 180
    ready = None
    expected_size = (ROOT / 'dist/app_payload.zip').stat().st_size
    while time.monotonic() < deadline:
        if process.poll() is not None:
            output, error = process.communicate()
            raise AssertionError(f'Setup exited during startup ({process.returncode}): {error!r}')
        bundles = [p for p in directory.glob('_MEI*') if p.is_dir()]
        if any((p / 'app_payload.zip').is_file()
               and (p / 'app_payload.zip').stat().st_size == expected_size
               and (p / 'resources/branding/favicon.png').is_file()
               and (p / 'app_payload.manifest.json').is_file() for p in bundles):
            ready = ready or time.monotonic()
            if time.monotonic() - ready >= 8:
                break
        time.sleep(.1)
    else:
        details = [(str(p.relative_to(directory)), p.stat().st_size) for p in directory.rglob('app_payload*') if p.is_file()]
        raise AssertionError(f'Setup required assets not observed within 180 seconds; expected payload size {expected_size}; observed {details}')
    print('PASS: exact Setup remains running with its extracted payload/resources (offscreen); Install was not pressed', flush=True)
finally:
    if process.poll() is None:
        stopped = subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'],
                                 capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
        assert stopped.returncode == 0, stopped.stderr
        process.wait(15)
    process.communicate()
    for attempt in range(20):
        try:
            temporary.cleanup()
            break
        except OSError as error:
            if getattr(error, 'winerror', None) not in (5, 32, 145) or attempt == 19:
                raise
            time.sleep(.25)
