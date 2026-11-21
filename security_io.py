"""Small shared file-safety helpers; no GUI or evidence writes."""
import csv
import io
import json
import os
import tempfile
from contextvars import ContextVar
from contextlib import contextmanager
from pathlib import Path


class ReportSaveCancelled(Exception):
    """A user declined replacing an existing report."""


_report_policy = ContextVar('findex_report_policy', default=None)


def report_identity(path):
    info = Path(path).lstat()
    return (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns)


def report_key(path):
    return os.path.normcase(str(Path(path).absolute()))


@contextmanager
def report_write_scope(authorize=None, protected_paths=(), protected_roots=(), approved=None):
    """Scope explicit replacement consent to one GUI export or inventory job."""
    policy = dict(authorize=authorize, paths=[Path(p).resolve() for p in protected_paths if p],
                  roots=[Path(p).resolve() for p in protected_roots if p],
                  approved=dict(approved or {}), written={})
    token = _report_policy.set(policy)
    try:
        yield policy
    finally:
        _report_policy.reset(token)


def check_report_target(target, policy):
    if linked(target) or not target.is_file():
        raise ValueError('Choose a regular report file; linked paths and folders cannot be replaced')
    if policy:
        resolved = target.resolve()
        if any(resolved == source or os.path.samefile(target, source) for source in policy['paths'] if source.exists()):
            raise ValueError('Source evidence cannot be replaced with a report')
        if any(resolved.is_relative_to(root) for root in policy['roots']):
            raise ValueError('Existing files inside an evidence folder cannot be replaced with reports')
    return report_identity(target)


@contextmanager
def new_report(path, binary=False):
    """Create reports exclusively, or atomically replace only with explicit consent."""
    target = Path(path).absolute()
    if str(target).lower().startswith(('\\\\.\\', '\\\\?\\globalroot\\')):
        raise ValueError('Device namespaces are not report destinations')
    # Windows alternate data streams can modify an existing evidence file even
    # when a stream name does not exist yet. Device paths are not report files.
    if any(':' in part or part.endswith((' ', '.')) or Path(part).is_reserved()
           for part in target.parts[1:]):
        raise ValueError('Choose a regular report filename without streams or reserved device names')
    target.parent.mkdir(parents=True, exist_ok=True)
    policy = _report_policy.get()
    key = report_key(target)
    replacement = target.exists() or target.is_symlink()
    temporary = None
    original = None
    if replacement:
        original = check_report_target(target, policy)
        if not policy:
            raise FileExistsError(str(target))
        allowed = policy['written'].get(key) or policy['approved'].get(key)
        if allowed is None:
            if not policy['authorize']:
                raise FileExistsError(str(target))
            if not policy['authorize'](target):
                raise ReportSaveCancelled()
        elif allowed != original:
            raise OSError('The report changed after replacement was approved. Please save again.')
        if check_report_target(target, policy) != original:
            raise OSError('The destination changed during confirmation. Please save again.')
        descriptor, name = tempfile.mkstemp(dir=target.parent, prefix='.findex-report-')
        temporary = Path(name)
        stream = os.fdopen(descriptor, 'wb' if binary else 'w', **({} if binary else {'encoding': 'utf-8', 'newline': ''}))
    else:
        # Exclusive creation also rejects aliases that appear after the existence check.
        stream = target.open('xb' if binary else 'x', **({} if binary else {'encoding': 'utf-8', 'newline': ''}))
    created = os.fstat(stream.fileno())
    try:
        with stream:
            yield stream
        if temporary:
            if check_report_target(target, policy) != original:
                raise OSError('The destination changed while saving. The previous file was preserved.')
            os.replace(temporary, target)
        if policy:
            policy['written'][key] = report_identity(target)
    except BaseException:
        if not temporary:
            try:
                current = target.lstat()
                if (current.st_dev, current.st_ino) == (created.st_dev, created.st_ino):
                    target.unlink()
            except OSError:
                pass
        raise
    finally:
        if temporary:
            temporary.unlink(missing_ok=True)


def write_report(path, content):
    with new_report(path) as output:
        output.write(content)


def save_workbook(workbook, path):
    with new_report(path, binary=True) as output:
        workbook.save(output)


def script_string(value):
    """Encode an untrusted string as a JS literal inside an HTML script element."""
    return json.dumps(str(value), ensure_ascii=True).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')


def csv_text(rows):
    output = io.StringIO(newline='')
    writer = csv.writer(output)
    for row in rows:
        values = []
        for value in row:
            value = str(value)
            if value.lstrip().startswith(('=', '+', '-', '@', '\t', '\r', '\n')):
                value = "'" + value
            values.append(value)
        writer.writerow(values)
    return output.getvalue()


def linked(path):
    path = Path(path)
    return path.is_symlink() or (hasattr(path, 'is_junction') and path.is_junction())


def evidence_walk(root, onerror=None):
    """Walk an explicitly selected root, skipping nested linked paths."""
    def warn(path):
        if onerror:
            onerror(OSError('Skipped linked path: ' + str(path)))
    for current, dirs, files in os.walk(root, followlinks=False, onerror=onerror):
        for name in list(dirs):
            if linked(Path(current) / name):
                dirs.remove(name)
                warn(Path(current) / name)
        regular = []
        for name in files:
            path = Path(current) / name
            if linked(path):
                warn(path)
            else:
                regular.append(name)
        yield current, dirs, regular
