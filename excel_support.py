"""Keep openpyxl usable when an unrelated, optional NumPy install is incomplete."""
import importlib
import sys


def load_openpyxl():
    before = set(sys.modules)
    try:
        return importlib.import_module('openpyxl'), None
    except AttributeError as error:
        if getattr(getattr(error, 'obj', None), '__name__', None) != 'numpy':
            raise
        reason = str(error)
    # Discard only openpyxl modules created by this interrupted import. Preserve
    # pre-existing modules and restore NumPy immediately after the retry.
    for name in set(sys.modules) - before:
        if name == 'openpyxl' or name.startswith('openpyxl.'):
            sys.modules.pop(name, None)
    absent = object()
    numpy_module = sys.modules.get('numpy', absent)
    sys.modules['numpy'] = None
    try:
        return importlib.import_module('openpyxl'), reason
    finally:
        if numpy_module is absent:
            sys.modules.pop('numpy', None)
        else:
            sys.modules['numpy'] = numpy_module
