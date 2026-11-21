"""Record the exact downloaded Windows/Python wheel set for future release builds."""
import email
import hashlib
import zipfile
from pathlib import Path

root = Path(__file__).resolve().parents[1]
entries = []
for wheel in sorted((root / 'scratch/release-wheels').glob('*.whl')):
    with zipfile.ZipFile(wheel) as archive:
        metadata = next(name for name in archive.namelist() if name.endswith('.dist-info/METADATA'))
        message = email.message_from_bytes(archive.read(metadata))
    digest = hashlib.sha256(wheel.read_bytes()).hexdigest()
    entries.append((message['Name'].lower(), f"{message['Name']}=={message['Version']} --hash=sha256:{digest}"))
if not entries:
    raise SystemExit('Download the reviewed wheel set first')
(root / 'requirements-release.lock').write_text(
    '# Windows x64 / Python 3.12 wheels. Reviewed release dependency closure.\n'
    + '\n'.join(line for _, line in sorted(entries)) + '\n', encoding='utf-8')
print('Locked', len(entries), 'wheel hashes')
