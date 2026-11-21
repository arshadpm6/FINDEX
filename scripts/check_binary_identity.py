"""Verify FINDEX product/version, icon resources and standard-user execution level."""
import sys
from pathlib import Path
import pefile

for argument in sys.argv[1:]:
    path = Path(argument)
    image = pefile.PE(str(path))
    strings = {}
    for group in getattr(image, 'FileInfo', []):
        for entry in group:
            for table in getattr(entry, 'StringTable', []):
                strings.update({k.decode(): v.decode() for k, v in table.entries.items()})
    assert strings.get('ProductName') == 'FINDEX', strings
    assert strings.get('ProductVersion') == '3.0.1', strings
    assert strings.get('OriginalFilename') == path.name, strings
    manifests = []
    resource_types = set()
    for entry in image.DIRECTORY_ENTRY_RESOURCE.entries:
        resource_types.add(entry.id)
        if entry.id == 24:
            for name in entry.directory.entries:
                for language in name.directory.entries:
                    data = language.data.struct
                    manifests.append(image.get_data(data.OffsetToData, data.Size).decode('utf-8'))
    assert 14 in resource_types, 'Missing Windows icon resource'
    assert manifests and all('asInvoker' in item and 'requireAdministrator' not in item for item in manifests), manifests
    print('Verified FINDEX 3.0.1 identity, icon and asInvoker:', path)
