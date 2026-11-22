"""Keep the Windows core release free of optional native helpers and unused Qt modules."""
from pathlib import PurePosixPath

QT_LIBRARIES = {"qt6core.dll", "qt6gui.dll", "qt6widgets.dll", "qt6svg.dll", "qt6opengl.dll"}
QT_EXTENSIONS = {"qtcore.pyd", "qtgui.pyd", "qtwidgets.pyd", "qtsvg.pyd"}
PLUGIN_GROUPS = {"platforms", "styles", "imageformats", "iconengines"}

def core_binaries(entries):
    kept = []
    for entry in entries:
        name = entry[0].replace("\\", "/")
        path = PurePosixPath(name)
        lower = path.name.lower()
        if "/plugins/" in name:
            group = name.split("/plugins/", 1)[1].split("/", 1)[0].lower()
            if group not in PLUGIN_GROUPS or lower == "qpdf.dll":
                continue
        if lower.startswith("qt6") and lower.endswith(".dll") and lower not in QT_LIBRARIES:
            continue
        if name.startswith("PySide6/") and lower.endswith(".pyd") and lower not in QT_EXTENSIONS:
            continue
        if lower in {"icuuc.dll", "icuin.dll"} or (lower.startswith("icudt") and lower.endswith(".dll")):
            continue
        kept.append(entry)
    return kept


def core_datas(entries):
    return [entry for entry in entries if not (entry[0].replace('\\', '/').startswith('PySide6/') and any(part in entry[0].replace('\\', '/').lower().split('/') for part in ('qml', 'translations')))]
