"""Payload validation shared by the release packager and installer."""
import hashlib
import json
from pathlib import PureWindowsPath


def payload_name(name):
    path = PureWindowsPath(name)
    if (not name or path.is_absolute() or path.drive or path.root or '..' in path.parts
            or '\\' in name or any(':' in part or part.endswith((' ', '.')) or
                PureWindowsPath(part).is_reserved() for part in path.parts)
            or path.as_posix() != name.rstrip('/')):
        raise ValueError('Unsafe payload path: ' + name)
    return path


def payload_manifest(archive, cancel_event=None):
    files = {}
    seen = set()
    for info in archive.infolist():
        payload_name(info.filename)
        key = info.filename.rstrip('/').casefold()
        if key in seen:
            raise ValueError('Duplicate payload path: ' + info.filename)
        seen.add(key)
        if info.is_dir():
            continue
        if (info.external_attr >> 16) & 0o170000 == 0o120000:
            raise ValueError('Linked payload entry: ' + info.filename)
        if key in {'install_manifest.json', 'uninstall.ps1', 'uninstall.bat'}:
            raise ValueError('Reserved installer path: ' + info.filename)
        digest = hashlib.sha256()
        with archive.open(info) as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                if cancel_event and cancel_event.is_set():
                    raise InterruptedError('Installation cancelled before extraction.')
                digest.update(chunk)
        files[info.filename] = {'size': info.file_size, 'sha256': digest.hexdigest()}
    return {'version': '3.0.1', 'files': files}


def verify_payload(archive, manifest_path, cancel_event=None):
    expected = json.loads(manifest_path.read_text(encoding='utf-8'))
    actual = payload_manifest(archive, cancel_event)
    if actual != expected or 'FINDEX.exe' not in actual['files']:
        raise ValueError('Application payload failed integrity verification. Obtain a fresh installer.')
    return actual
