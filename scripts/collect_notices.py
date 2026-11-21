"""Copy license/notice files shipped in the installed, pinned Python wheels."""
import importlib.metadata
import shutil
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[1]
destination = root / 'resources/third_party'
count = 0
for dist in importlib.metadata.distributions():
    for entry in dist.files or []:
        if any(token in Path(entry).name.lower() for token in ('license', 'licence', 'copying', 'copyright', 'notice')):
            source = Path(dist.locate_file(entry))
            if source.is_file() and '..' not in entry.parts:
                target = destination / dist.metadata['Name'] / str(entry)
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
                count += 1
print('Copied', count, 'wheel license/notice files; corresponding-source obligations still require review')
python_license = Path(sys.base_prefix) / 'LICENSE.txt'
if python_license.is_file():
    target = destination / 'Python' / 'LICENSE.txt'
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(python_license, target)
else:
    raise SystemExit('Python license text is missing; obtain it for the exact runtime before packaging')
