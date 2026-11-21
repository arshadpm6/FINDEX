"""Refuse unreviewed Python packages in the release build environment."""
import importlib.metadata
import re
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[1]
normalize = lambda name: re.sub(r'[-_.]+', '-', name).lower()
expected = {}
for line in (root / 'requirements-release.lock').read_text(encoding='utf-8').splitlines():
    if line and not line.startswith('#'):
        name, version = line.split()[0].split('==')
        expected[normalize(name)] = version
actual = {normalize(dist.metadata['Name']): dist.version for dist in importlib.metadata.distributions()}
if actual != expected or sys.prefix == sys.base_prefix:
    raise SystemExit('Release environment differs from the lock. Use a clean isolated environment.\n'
                     + str({'expected': expected, 'actual': actual}))
print('Release environment matches all', len(expected), 'pinned dependencies')
