#!/usr/bin/env python3
"""
================================================================================
FINDEX — FILE INTELLIGENCE & EXAMINATION
================================================================================
Enterprise Desktop GUI Application built with Python & PySide6 (Qt)
Design Theme: Surface Modern Slate/Navy Dashboard

Features:
1. File Details Inventory:
   - Scans directory trees recursively
   - Computes file sizes, formats, creation/modification dates, paths
   - Extracts exact media durations via ffprobe
   - Computes cryptographic checksums (MD5, SHA-1, SHA-256)
   - Configurable column selection with audit presets
   - Numbering modes: Per-Folder (resets to 1 per folder) or Continuous
   - Extension filters: Audio, Video, Images, Documents, Archives, Other
   - Exports to styled Excel (.xlsx) & self-contained interactive HTML reports
2. Cryptographic Hash Verification:
   - Batch hashing (MD5, SHA-1, SHA-256, SHA-512 in single read pass)
   - Verification against expected checksum manifests (.sha256, .md5)
   - Real-time animated futuristic verification hero loader
   - Side-by-side two-file comparator
   - Detailed hash inspector with 1-click clipboard copy
   - Exports results to Checksum file, Excel, HTML, CSV

Author: Advanced Agentic Software Engineer
Version: 3.0.1
================================================================================
"""

import os
import sys
import math
import io
import json
import base64
import time
import html as html_module
import uuid
import shutil
import hashlib
import datetime
import subprocess
import threading
import copy
import functools
import queue
from pathlib import Path
from security_io import (new_report, write_report, save_workbook, script_string, csv_text,
                         evidence_walk, linked, report_write_scope, report_identity,
                         report_key, check_report_target, ReportSaveCancelled)

def get_app_root():
    """Returns the base application directory whether running as script or frozen PyInstaller executable."""
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            return Path(sys._MEIPASS).resolve()
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

APP_ROOT = get_app_root()

# Enable High DPI scaling on Windows
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QFrame, QScrollArea,
    QFileDialog, QLineEdit, QCheckBox, QRadioButton, QButtonGroup,
    QComboBox, QProgressBar, QStackedWidget, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QDialog, QTextEdit, QSplitter, QSizePolicy,
    QGraphicsDropShadowEffect, QMessageBox, QSpacerItem, QGroupBox,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QMenu, QInputDialog
)
from PySide6.QtCore import (
    Qt, Signal, Slot, QThread, QTimer, QPropertyAnimation,
    QEasingCurve, QSize, Property, QRect, QRectF, QPoint, QPointF, QUrl, QSettings, QByteArray
)
from PySide6.QtGui import (
    QFont, QColor, QPainter, QPainterPath, QLinearGradient,
    QRadialGradient, QConicalGradient, QPen, QBrush, QIcon, QPixmap,
    QFontMetrics, QAction, QDesktopServices, QCursor, QPalette
)
from PySide6.QtSvg import QSvgRenderer
from dialog_theme import install_message_box_theme_filter
from report_branding import (report_logo_uri, report_icon, report_brand_header,
                             report_head_tags, REPORT_BRAND_CSS)

from excel_support import load_openpyxl

OPENPYXL_IMPORT_ERROR = None
OPENPYXL_NUMPY_FALLBACK = None
try:
    openpyxl, OPENPYXL_NUMPY_FALLBACK = load_openpyxl()
    from openpyxl.styles import Font as XlFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except (ImportError, AttributeError) as exc:
    if isinstance(exc, AttributeError) and getattr(
        getattr(exc, "obj", None), "__name__", None
    ) != "numpy":
        raise
    OPENPYXL_AVAILABLE = False
    OPENPYXL_IMPORT_ERROR = exc


def _openpyxl_unavailable_error(message):
    if OPENPYXL_IMPORT_ERROR is not None:
        message = f"{message} Import failed: {OPENPYXL_IMPORT_ERROR}"
    return RuntimeError(message)


# ==============================================================================
# CONSTANTS & THEME PALETTE (Surface Modern Slate / Deep Navy)
# ==============================================================================

APP_NAME = "FINDEX"
APP_SUBTITLE = "File Intelligence & Examination"
APP_VERSION = "3.0.1"
APP_TAGLINE = "Know Every File."
APP_WINDOW_TITLE = f"{APP_NAME} – {APP_SUBTITLE} v{APP_VERSION}"
MIN_WINDOW_WIDTH = 960
MIN_WINDOW_HEIGHT = 600

# Surface Dashboard Color Tokens (matching reference image)
COLOR_BG_DEEP       = "#0b0e17"    # Deep space window backdrop
COLOR_BG_DARK       = "#111422"    # Secondary dark surface
COLOR_CARD_BG       = "#161b2b"    # Modern card background
COLOR_CARD_INNER    = "#1c2236"    # Nested card / stat block
COLOR_CARD_BORDER   = "#232b43"    # Subtle glowing card border
COLOR_CARD_HOVER    = "#1e253c"    # Card hover state

COLOR_ACCENT        = "#3b82f6"    # Surface vibrant blue
COLOR_ACCENT_HOVER  = "#60a5fa"    # Hover lighter blue
COLOR_ACCENT_DARK   = "#2563eb"    # Active pressed blue
COLOR_ACCENT_GLOW   = "rgba(59, 130, 246, 0.25)"

COLOR_SUCCESS       = "#2f7d46"    # Standard verification green
COLOR_SUCCESS_BG    = "rgba(128, 204, 147, 0.15)"
COLOR_SUCCESS_BORDER= "#287a45"

COLOR_DANGER        = "#ef4444"    # Mismatch red
COLOR_DANGER_BG     = "rgba(239, 68, 68, 0.15)"
COLOR_DANGER_BORDER = "#dc2626"

COLOR_WARNING       = "#f59e0b"    # Amber warning
COLOR_CYAN          = "#06b6d4"    # Cyan accent
COLOR_PURPLE        = "#8b5cf6"    # Purple accent

COLOR_TEXT_PRIMARY   = "#f1f5f9"   # Bright crisp text
COLOR_TEXT_SECONDARY = "#8e98b0"   # Muted metadata
COLOR_TEXT_MUTED     = "#55607c"   # Disabled / quiet text
COLOR_INPUT_BG       = "#0e121d"   # Dark input background
COLOR_INPUT_BORDER   = "#232b43"

# Excel Export Hex Colors
XL_COLOR_TITLE_BG   = "111422"
XL_COLOR_FOLDER_BG  = "3B82F6"
XL_COLOR_HEADER_BG  = "1E293B"
XL_COLOR_STRIPE_ODD = "FFFFFF"
XL_COLOR_STRIPE_EVEN= "F8FAFC"
XL_COLOR_SI_BG      = "EFF6FF"
XL_COLOR_BORDER     = "E2E8F0"
XL_COLOR_WHITE      = "FFFFFF"
XL_COLOR_SUCCESS_BG = "D1FAE5"
XL_COLOR_SUCCESS_TXT= "065F46"
XL_COLOR_DANGER_BG  = "FEE2E2"
XL_COLOR_DANGER_TXT = "991B1B"
XL_COLOR_INFO_BG    = "DBEAFE"
XL_COLOR_INFO_TXT   = "1E40AF"

# File Extension Categories
AUDIO_EXTENSIONS = {
    ".mp3", ".wav", ".ogg", ".opus", ".flac", ".aac",
    ".m4a", ".wma", ".aiff", ".aif", ".amr", ".au"
}
VIDEO_EXTENSIONS = {
    ".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv",
    ".webm", ".mpeg", ".mpg", ".m4v", ".3gp", ".ts",
    ".mts", ".m2ts", ".vob", ".asf"
}
IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff",
    ".tif", ".webp", ".svg", ".ico", ".raw", ".cr2", ".nef", ".dng", ".arw", ".heic"
}
DOCUMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt",
    ".pptx", ".txt", ".rtf", ".odt", ".ods", ".odp",
    ".csv", ".md"
}
ARCHIVE_EXTENSIONS = {
    ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2",
    ".xz", ".iso", ".cab"
}

MEDIA_EXTENSIONS = AUDIO_EXTENSIONS | VIDEO_EXTENSIONS
ALL_NAMED_EXTENSIONS = (
    AUDIO_EXTENSIONS | VIDEO_EXTENSIONS | IMAGE_EXTENSIONS |
    DOCUMENT_EXTENSIONS | ARCHIVE_EXTENSIONS
)

HASH_ALGORITHMS = ["SHA-256", "MD5", "SHA-1", "SHA-512"]

ALL_COLUMNS = [
    {"id": "si", "name": "SI No.", "width": 8, "align": "center", "locked": True, "default": True},
    {"id": "name", "name": "File Name", "width": 50, "align": "left", "locked": True, "default": True},
    {"id": "size", "name": "File Size", "width": 14, "align": "center", "locked": False, "default": True},
    {"id": "format", "name": "Format", "width": 12, "align": "center", "locked": False, "default": True},
    {"id": "duration", "name": "Duration", "width": 12, "align": "center", "locked": False, "default": True},
    {"id": "modified", "name": "Modified Date", "width": 20, "align": "center", "locked": False, "default": True},
    {"id": "created", "name": "Created Date", "width": 20, "align": "center", "locked": False, "default": False},
    {"id": "path", "name": "Full Path", "width": 60, "align": "left", "locked": False, "default": False},
    {"id": "md5", "name": "MD5", "width": 36, "align": "center", "locked": False, "default": False},
    {"id": "sha1", "name": "SHA-1", "width": 44, "align": "center", "locked": False, "default": False},
    {"id": "sha256", "name": "SHA-256", "width": 68, "align": "center", "locked": False, "default": False},
]
COLUMN_MAP = {col["id"]: col for col in ALL_COLUMNS}


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================


HASH_CHUNK_SIZE = 65536
ENABLED_HASH_ALGORITHMS = set(HASH_ALGORITHMS)


def configure_hash_preferences(settings):
    global HASH_CHUNK_SIZE, ENABLED_HASH_ALGORITHMS
    sizes = [65536, 131072, 262144, 1048576, 4194304]
    index = settings.value("chunk_size_index", 1, type=int)
    HASH_CHUNK_SIZE = sizes[index] if index in range(len(sizes)) else sizes[1]
    ENABLED_HASH_ALGORITHMS = {"SHA-256"}
    for algorithm in ("MD5", "SHA-1", "SHA-512"):
        if settings.value("hash_enabled_" + algorithm, True, type=bool):
            ENABLED_HASH_ALGORITHMS.add(algorithm)


def metadata_report_basename(record):
    source = str(Path(record.get("file_path") or record.get("file_name") or "evidence").resolve())
    stem = Path(record.get("file_name") or source).stem
    safe = "".join(c for c in stem if c.isalnum() or c in (" ", "_", "-")).strip()[:80] or "evidence"
    return safe + "_" + hashlib.sha256(os.path.normcase(source).encode("utf-8")).hexdigest()[:16]


def integrity_status_text(info):
    if not info or info.get("verified") is None:
        return "Not compared"
    if info.get("status") == "HASH_FAILED" or not info.get("pre") or not info.get("post"):
        return "Hash calculation failed"
    return "Verified Match" if info.get("verified") is True else "Mismatch Detected"


def set_excel_literal(cell, value):
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"


def report_export_errors(function):
    """Keep filesystem failures inside the UI boundary, including partial batches."""
    @functools.wraps(function)
    def export(self, *args, **kwargs):
        policy = None
        try:
            paths, roots = self._report_evidence_targets()
            with report_write_scope(authorize=self._confirm_report_replace,
                                    protected_paths=paths, protected_roots=roots) as policy:
                return function(self, *args, **kwargs)
        except ReportSaveCancelled:
            if policy and policy['written']:
                QMessageBox.information(self, 'Export Stopped',
                    'The remaining export was cancelled. Reports already saved were retained.')
            return
        except FileExistsError:
            QMessageBox.warning(self, "Report Already Exists", "The existing file was preserved. Choose a new report name or output folder.")
        except (OSError, ValueError, RuntimeError) as error:
            QMessageBox.critical(self, "Export Failed",
                f"The export could not be completed.\n\n{error}\n\n"
                "Some output files may already have been saved. Check the selected destination before retrying.")
    return export


def application_icon():
    """FINDEX desktop icon derived only from the primary favicon artwork."""
    branding = APP_ROOT / "resources" / "branding"
    desktop_icon = branding / "icon.ico"
    return QIcon(str(desktop_icon if desktop_icon.is_file() else branding / "favicon.png"))


def format_size(size_bytes):
    """Convert integer byte count to human-readable size string."""
    if size_bytes is None:
        return "0 B"
    size = float(size_bytes)
    for unit in ["B", "KB", "MB", "GB", "TB", "PB"]:
        if size < 1024.0 or unit == "PB":
            return f"{size:.2f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024.0
    return f"{size:.2f} PB"


def format_inventory_size(size_bytes, size_unit="auto"):
    """Apply the File Inventory report preference; retain raw bytes separately."""
    unit = str(size_unit).lower()
    divisor = {"kb": 1024, "mb": 1024 ** 2}.get(unit)
    if divisor is None:
        return format_size(size_bytes)
    return f"{float(size_bytes or 0) / divisor:.2f} {unit.upper()}"


def inventory_report_columns(enabled_column_ids, size_unit="auto"):
    columns = [dict(COLUMN_MAP[cid]) for cid in enabled_column_ids if cid in COLUMN_MAP]
    if not columns:
        columns = [dict(COLUMN_MAP["si"]), dict(COLUMN_MAP["name"])]
    if str(size_unit).lower() in ("kb", "mb"):
        for column in columns:
            if column["id"] == "size":
                column["name"] = f"File Size ({str(size_unit).upper()})"
    return columns


def format_duration(seconds):
    """Convert float or integer seconds to HH:MM:SS format."""
    if seconds is None:
        return "—"
    try:
        total_seconds = int(round(float(seconds)))
        if total_seconds < 0:
            return "—"
        hours, remainder = divmod(total_seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    except (ValueError, TypeError):
        return "—"


def truncate_path(path_str, max_len=60):
    """Truncate path string with ellipsis if too long."""
    if len(path_str) <= max_len:
        return path_str
    part = (max_len - 3) // 2
    return f"{path_str[:part]}...{path_str[-part:]}"


def find_ffprobe_executable():
    """Locate ffprobe on the system, prioritizing hardcoded path."""
    common_paths = [
        str(APP_ROOT / "external" / "ffmpeg" / "ffprobe.exe"),
        str(APP_ROOT / "external" / "ffprobe.exe"),
        r"C:\ffmpeg\bin\ffprobe.exe",
        r"C:\Program Files\ffmpeg\bin\ffprobe.exe",
    ]
    for p in common_paths:
        if os.path.isfile(p):
            return p
    probe = shutil.which("ffprobe")
    if probe:
        return probe
    return None

FFPROBE_BIN = find_ffprobe_executable()


def extract_media_duration(file_path):
    """Query media duration using ffprobe with hidden console."""
    if not FFPROBE_BIN:
        return None
    try:
        startupinfo = None
        creationflags = 0
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW
        cmd = [
            FFPROBE_BIN, "-v", "error", "-protocol_whitelist", "file,pipe", "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1", str(file_path)
        ]
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=15,
            check=True, shell=False, startupinfo=startupinfo, creationflags=creationflags
        )
        val = result.stdout.strip()
        if val:
            sec = float(val)
            if sec >= 0:
                return sec
    except Exception:
        pass
    return None


def compute_all_file_hashes(file_path, cancel_event=None, chunk_callback=None, algorithms=None, chunk_size=None):
    """Calculate selected hashes in one cancellable streaming pass."""
    selected = set(algorithms) if algorithms is not None else set(ENABLED_HASH_ALGORITHMS)
    constructors = {"MD5": hashlib.md5, "SHA-1": hashlib.sha1,
                    "SHA-256": hashlib.sha256, "SHA-512": hashlib.sha512}
    if not selected or selected - constructors.keys():
        raise ValueError("Select supported hash algorithms")
    digests = {name: constructors[name]() for name in selected}
    if chunk_size is not None and chunk_size <= 0:
        raise ValueError("Hash chunk size must be positive")
    total_read = 0
    with open(file_path, "rb") as stream:
        before = os.fstat(stream.fileno())
        while True:
            if cancel_event and cancel_event.is_set():
                return None
            chunk = stream.read(chunk_size or HASH_CHUNK_SIZE)
            if not chunk:
                break
            for digest in digests.values():
                digest.update(chunk)
            total_read += len(chunk)
            if chunk_callback:
                chunk_callback(total_read)
        after = os.fstat(stream.fileno())
        current = os.stat(file_path)
        # Windows handle and path APIs can expose different creation-time values.
        signature = lambda st: (st.st_dev, st.st_ino, st.st_size, st.st_mtime_ns)
        if signature(before) != signature(after) or signature(after) != signature(current) or total_read != after.st_size:
            raise OSError("File changed while hashing; no checksum was accepted. Retry with stable evidence.")
    return {name: digest.hexdigest() for name, digest in digests.items()}


def detect_hash_algorithm(hash_str):
    """Detect hash algorithm from hex character length."""
    clean = hash_str.strip().lower()
    length = len(clean)
    if not clean or any(c not in "0123456789abcdef" for c in clean): return None
    if length == 32: return "MD5"
    elif length == 40: return "SHA-1"
    elif length == 64: return "SHA-256"
    elif length == 128: return "SHA-512"
    return None


def open_evidence_file(parent, path):
    """Keep external viewing, with explicit consent for executable associations."""
    target = Path(path)
    if not target.is_file():
        QMessageBox.warning(parent, "Open File", "The selected file is no longer available.")
        return False
    executable = {".exe", ".com", ".bat", ".cmd", ".ps1", ".psm1", ".vbs", ".vbe",
                  ".js", ".jse", ".wsf", ".wsh", ".hta", ".scr", ".msi", ".msp",
                  ".cpl", ".lnk", ".url", ".reg", ".py", ".pyw", ".jar", ".application",
                  ".html", ".htm", ".svg", ".mht", ".mhtml"}
    if target.suffix.lower() in executable:
        answer = QMessageBox.question(parent, "Open Potentially Executable File",
            "Windows may execute this file or its target. FINDEX will not run it for examination.\n\n"
            "Open it externally anyway?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return False
    return QDesktopServices.openUrl(QUrl.fromLocalFile(str(target.resolve())))


# ==============================================================================
# SECTION 3: EXIFTOOL METADATA EXTRACTION ENGINE & INTEGRITY ADAPTER
# ==============================================================================

METADATA_CATEGORIES = [
    ("File Information", ["File", "System"]),
    ("EXIF", ["EXIF", "IFD0", "ExifIFD", "SubIFD", "InteropIFD"]),
    ("GPS", ["GPS"]),
    ("XMP", ["XMP"]),
    ("IPTC", ["IPTC", "Photoshop"]),
    ("MakerNotes", ["MakerNotes", "Canon", "Nikon", "Sony", "Apple", "Fujifilm", "Olympus", "Panasonic", "Pentax", "Leica"]),
    ("ICC Profile", ["ICC_Profile", "ColorSpace"]),
    ("Image Information", ["PNG", "JPEG", "GIF", "TIFF", "BMP", "Composite", "JUMBF", "C2PA"]),
    ("Video Metadata", ["QuickTime", "MP4", "Matroska", "RIFF", "AVI", "MPEG", "Track"]),
    ("Audio Metadata", ["ID3", "Vorbis", "FLAC", "Audio", "M4A", "AIFF"]),
    ("Document Metadata", ["PDF", "PostScript", "Flash", "DOC", "XLS", "PPT"]),
    ("Other Metadata", []),
]


class ExifToolEngine:
    """
    Subprocess-based interface for the official Phil Harvey ExifTool executable.
    Does not use python EXIF parser replacements. Dynamically detects actual ExifTool version.
    """
    _cached_version = None
    _cached_path = None

    @classmethod
    def find_executable(cls, custom_path=None):
        if custom_path and os.path.isfile(custom_path):
            return str(Path(custom_path).resolve())

        # 1. Bundled / Application local external/exiftool directory
        local_cand = APP_ROOT / "external" / "exiftool" / "exiftool.exe"
        if local_cand.exists():
            return str(local_cand.resolve())

        # 2. Check system PATH
        which_path = shutil.which("exiftool") or shutil.which("exiftool.exe")
        if which_path:
            return str(Path(which_path).resolve())

        # 3. Common Windows install locations
        win_candidates = [
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\ExifToolGUI\exiftool.exe"),
            os.path.expandvars(r"%PROGRAMFILES%\ExifTool\exiftool.exe"),
            os.path.expandvars(r"%PROGRAMFILES(X86)%\ExifTool\exiftool.exe"),
            r"C:\ExifTool\exiftool.exe",
        ]
        for c in win_candidates:
            if os.path.isfile(c):
                return str(Path(c).resolve())

        return None

    @classmethod
    def command(cls, exe, arguments):
        # ExifTool configuration files contain executable Perl. Never auto-load them.
        arguments = ["-config", "", *arguments]
        # Avoid the console-launching Windows wrapper for the bundled engine.
        runtime = Path(exe).parent / "exiftool_files"
        perl, script = runtime / "perl.exe", runtime / "exiftool.pl"
        if Path(exe).name.lower() == "exiftool.exe" and perl.is_file() and script.is_file():
            return [str(perl), "-I", "lib", "exiftool.pl", *arguments]
        return [exe, *arguments]

    @classmethod
    def working_directory(cls, exe):
        runtime = Path(exe).parent / "exiftool_files"
        if Path(exe).name.lower() == "exiftool.exe" and (runtime / 'perl.exe').is_file() and (runtime / 'exiftool.pl').is_file():
            return str(runtime)
        return str(Path(exe).parent)

    @staticmethod
    def environment():
        environment = os.environ.copy()
        environment.update(LC_ALL='C', LC_CTYPE='C', LANG='C', PERL5OPT='', PERL5LIB='')
        return environment

    @classmethod
    def get_version(cls, custom_path=None):
        """Execute 'exiftool.exe -ver' and capture actual detected version dynamically."""
        exe = cls.find_executable(custom_path)
        if not exe:
            return None, False, "ExifTool executable not found. Please place exiftool.exe in external/exiftool/ or set path."

        try:
            startupinfo = None
            creationflags = 0
            if sys.platform == "win32":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                creationflags = subprocess.CREATE_NO_WINDOW

            res = subprocess.run(
                cls.command(exe, ["-ver"]),
                cwd=cls.working_directory(exe),
                env=cls.environment(),
                capture_output=True,
                text=True,
                timeout=8,
                shell=False,
                startupinfo=startupinfo,
                creationflags=creationflags
            )
            ver = res.stdout.strip()
            if res.returncode == 0 and ver and all(c in "0123456789." for c in ver):
                cls._cached_version = ver
                cls._cached_path = exe
                return ver, True, exe
        except Exception as ex:
            return None, False, str(ex)

        return None, False, "Unable to obtain ExifTool version output."

    @classmethod
    def extract_metadata(cls, file_path, custom_path=None, timeout=45):
        """
        Extracts metadata using official ExifTool executable with safe argument list:
        exiftool.exe -json -G -a -s -struct "filename"
        Preserves raw output, parses duplicate/grouped tags safely, and categorizes into logical groups.
        """
        exe = cls.find_executable(custom_path)
        if not exe:
            return {
                "success": False,
                "error": "ExifTool executable not found.",
                "file_path": str(file_path),
                "raw_json": "",
                "parsed": {},
                "categories": {},
                "warnings": ["ExifTool executable missing."]
            }

        p = Path(file_path)
        try:
            is_regular_file = p.is_file()
            file_size = p.stat().st_size if is_regular_file else 0
        except OSError:
            is_regular_file, file_size = False, 0
        if not is_regular_file:
            return {
                "success": False,
                "error": f"File does not exist: {file_path}",
                "file_path": str(file_path),
                "raw_json": "",
                "parsed": {},
                "categories": {},
                "warnings": ["File not found on disk."]
            }

        # A UTF-8 stdin argument file avoids Windows ANSI argv loss for Unicode paths.
        cmd = cls.command(exe, ["-charset", "filename=UTF8", "-json", "-G", "-a", "-s", "-struct", "-@", "-"])
        startupinfo = None
        creationflags = 0
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW

        try:
            proc = subprocess.run(
                cmd,
                input=str(p.resolve()) + "\n",
                cwd=cls.working_directory(exe),
                env=cls.environment(),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                shell=False,
                startupinfo=startupinfo,
                creationflags=creationflags
            )
            raw_stdout = proc.stdout
            raw_stderr = proc.stderr
            warnings = []
            if raw_stderr:
                for line in raw_stderr.splitlines():
                    line_s = line.strip()
                    if line_s:
                        warnings.append(line_s)

            if not raw_stdout.strip():
                return {
                    "success": False,
                    "error": f"ExifTool returned empty output (Code {proc.returncode}). {raw_stderr.strip()}",
                    "file_path": str(file_path),
                    "file_name": p.name,
                    "file_size": file_size,
                    "raw_json": "",
                    "parsed": {},
                    "categories": {},
                    "warnings": warnings,
                    "exiftool_version": cls._cached_version or "Unknown"
                }

            parsed_list = json.loads(raw_stdout)
            if not isinstance(parsed_list, list) or len(parsed_list) != 1 or not isinstance(parsed_list[0], dict):
                raise ValueError("ExifTool returned an invalid metadata record")
            record = parsed_list[0]
            engine_errors = [str(v) for k, v in record.items() if k.rsplit(":", 1)[-1] == "Error"]
            warnings.extend(str(v) for k, v in record.items() if k.rsplit(":", 1)[-1] == "Warning")
            extraction_error = "; ".join(engine_errors) or (f"ExifTool exited with code {proc.returncode}: {raw_stderr.strip()}" if proc.returncode else None)

            categories = cls.group_metadata(record)

            return {
                "success": extraction_error is None,
                "error": extraction_error,
                "file_path": str(p.resolve()),
                "file_name": p.name,
                "file_size": file_size,
                "raw_json": raw_stdout,
                "parsed": record,
                "categories": categories,
                "warnings": warnings,
                "exiftool_version": cls._cached_version or "Unknown"
            }
        except subprocess.TimeoutExpired:
            return {
                "success": False,
                "error": f"ExifTool extraction timed out after {timeout} seconds.",
                "file_path": str(file_path),
                "file_name": p.name,
                "file_size": file_size,
                "raw_json": "",
                "parsed": {},
                "categories": {},
                "warnings": ["Timeout expired."],
                "exiftool_version": cls._cached_version or "Unknown"
            }
        except Exception as ex:
            return {
                "success": False,
                "error": str(ex),
                "file_path": str(file_path),
                "file_name": p.name,
                "file_size": file_size,
                "raw_json": "",
                "parsed": {},
                "categories": {},
                "warnings": [str(ex)],
                "exiftool_version": cls._cached_version or "Unknown"
            }

    @classmethod
    def group_metadata(cls, parsed_record):
        """Organize raw parsed ExifTool tags into the 12 logical categories."""
        groups = {cat_name: [] for cat_name, _ in METADATA_CATEGORIES}

        for full_key, value in parsed_record.items():
            if full_key == "SourceFile":
                continue

            if ":" in full_key:
                parts = full_key.split(":")
                group_prefix = parts[0]
                tag_name = ":".join(parts[1:])
            else:
                group_prefix, tag_name = "Other", full_key

            assigned = False
            for cat_name, prefixes in METADATA_CATEGORIES:
                if prefixes:
                    for pfx in prefixes:
                        if group_prefix.lower() == pfx.lower() or group_prefix.lower().startswith(pfx.lower()):
                            groups[cat_name].append({
                                "tag": tag_name,
                                "raw_key": full_key,
                                "group": group_prefix,
                                "value": value
                            })
                            assigned = True
                            break
                    if assigned:
                        break

            if not assigned:
                groups["Other Metadata"].append({
                    "tag": tag_name,
                    "raw_key": full_key,
                    "group": group_prefix,
                    "value": value
                })

        return {k: v for k, v in groups.items() if len(v) > 0}


class MetadataHashAdapter:
    """
    Forensic integrity adapter reusing the application's existing compute_all_file_hashes engine.
    Calculates cryptographic hashes before and after metadata extraction to verify integrity.
    """
    @staticmethod
    def calculate_integrity(file_path, cancel_event=None):
        return compute_all_file_hashes(file_path, cancel_event=cancel_event)

    @staticmethod
    def verify_integrity(pre_hashes, file_path, cancel_event=None):
        post_hashes = compute_all_file_hashes(file_path, cancel_event=cancel_event, algorithms=pre_hashes.keys() if pre_hashes else None)
        if not pre_hashes or not post_hashes:
            return {
                "verified": False,
                "status": "HASH_FAILED",
                "pre": pre_hashes or {},
                "post": post_hashes or {},
                "message": "Hash calculation failed for file."
            }

        matches = "SHA-256" in pre_hashes and pre_hashes == post_hashes
        return {
            "verified": matches,
            "status": "VERIFIED_PERFECT_MATCH" if matches else "INTEGRITY_MISMATCH",
            "pre": pre_hashes,
            "post": post_hashes,
            "message": "Pre and post-extraction hashes match. This does not establish authenticity or provenance." if matches else "Pre and post hashes do not match! Possible file alteration."
        }


# ==============================================================================
# SECTION 4: FILE INVENTORY REPORT GENERATORS
# ==============================================================================

class ExcelReportGenerator:
    """Creates styled Excel reports for folder tree inventory."""

    @staticmethod
    def generate(output_file, folder_data, enabled_column_ids, main_folder_path, total_files, total_size, size_unit="auto"):
        if not OPENPYXL_AVAILABLE:
            raise _openpyxl_unavailable_error(
                "The 'openpyxl' package is required to generate Excel reports."
            ) from OPENPYXL_IMPORT_ERROR

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "File Inventory"
        ws.sheet_view.showGridLines = False

        active_cols = inventory_report_columns(enabled_column_ids, size_unit)
        num_cols = len(active_cols) if active_cols else 2

        thin_side = Side(style="thin", color=XL_COLOR_BORDER)
        data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        # Row 1: Title
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        folder_display = Path(main_folder_path).name or str(main_folder_path)
        title_cell = ws.cell(row=1, column=1)
        title_cell.value = f"FILE INVENTORY REPORT — {folder_display}"
        title_cell.font = XlFont(name="Segoe UI", size=15, bold=True, color=XL_COLOR_WHITE)
        title_cell.fill = PatternFill("solid", fgColor=XL_COLOR_TITLE_BG)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        # Row 2: Metadata
        gen_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        meta_str = (
            f"Source folder: {main_folder_path}    |    "
            f"Generated: {gen_time}    |    "
            f"Total files: {total_files:,} ({format_inventory_size(total_size, size_unit)})"
        )
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        meta_cell = ws.cell(row=2, column=1, value=meta_str)
        meta_cell.font = XlFont(name="Segoe UI", size=9, bold=True, color=XL_COLOR_HEADER_BG)
        meta_cell.fill = PatternFill("solid", fgColor=XL_COLOR_SI_BG)
        meta_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22

        # Row 3: Blank separator
        ws.row_dimensions[3].height = 12

        # Rows 4+: Folder sections
        current_row = 4
        first_data_row = None

        for folder in folder_data:
            files_in_folder = folder["files"]
            folder_label = folder["folder"]
            folder_size_str = format_inventory_size(folder["total_size"], size_unit)
            plural = "file" if len(files_in_folder) == 1 else "files"

            # Folder Header
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
            f_cell = ws.cell(row=current_row, column=1)
            f_cell.value = f"FOLDER: {folder_label}   ({len(files_in_folder)} {plural}, {folder_size_str})"
            f_cell.font = XlFont(name="Segoe UI", size=11, bold=True, color=XL_COLOR_WHITE)
            f_cell.fill = PatternFill("solid", fgColor=XL_COLOR_FOLDER_BG)
            f_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[current_row].height = 24
            current_row += 1

            # Column Header
            ws.row_dimensions[current_row].height = 22
            for col_idx, col_def in enumerate(active_cols, start=1):
                c = ws.cell(row=current_row, column=col_idx, value=col_def["name"])
                c.font = XlFont(name="Segoe UI", size=10, bold=True, color=XL_COLOR_WHITE)
                c.fill = PatternFill("solid", fgColor=XL_COLOR_HEADER_BG)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = data_border
            current_row += 1

            # Data rows
            if not files_in_folder:
                ws.row_dimensions[current_row].height = 20
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=num_cols)
                empty_cell = ws.cell(row=current_row, column=1, value="No matching files found in this folder")
                empty_cell.font = XlFont(name="Segoe UI", size=9, italic=True, color="64748B")
                empty_cell.alignment = Alignment(horizontal="center", vertical="center")
                empty_cell.border = data_border
                current_row += 1
            else:
                for file_idx, f_info in enumerate(files_in_folder):
                    if first_data_row is None:
                        first_data_row = current_row

                    ws.row_dimensions[current_row].height = 20
                    is_even = (file_idx % 2 == 1)
                    stripe_bg = XL_COLOR_STRIPE_EVEN if is_even else XL_COLOR_STRIPE_ODD

                    for col_idx, col_def in enumerate(active_cols, start=1):
                        cid = col_def["id"]
                        val = f_info.get(cid, "—")
                        if cid == "size" and "size_bytes" in f_info:
                            val = format_inventory_size(f_info["size_bytes"], size_unit)
                        c = ws.cell(row=current_row, column=col_idx)
                        set_excel_literal(c, val)
                        c.font = XlFont(name="Segoe UI", size=9)
                        c.border = data_border

                        if cid == "si":
                            c.fill = PatternFill("solid", fgColor=XL_COLOR_SI_BG)
                            c.font = XlFont(name="Segoe UI", size=9, bold=True, color=XL_COLOR_HEADER_BG)
                        else:
                            c.fill = PatternFill("solid", fgColor=stripe_bg)

                        if col_def["align"] == "center":
                            c.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            c.alignment = Alignment(horizontal="left", vertical="center")

                    current_row += 1

            # Blank row between folders
            ws.row_dimensions[current_row].height = 12
            current_row += 1

        # Set Column Widths
        for col_idx, col_def in enumerate(active_cols, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = col_def["width"]

        # Freeze Panes
        ws.freeze_panes = f"A{first_data_row if first_data_row else 6}"

        # Save workbook
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        save_workbook(wb, output_file)


def get_logo_favicon_b64():
    """Use the dedicated FINDEX report asset; never reuse the legacy icon cache."""
    return report_logo_uri(APP_ROOT)


class HtmlReportGenerator:
    """Generates self-contained interactive HTML inventory reports."""

    @staticmethod
    def generate(output_file, folder_data, enabled_column_ids, main_folder_path, total_files, total_size, size_unit="auto"):
        active_cols = inventory_report_columns(enabled_column_ids, size_unit)

        gen_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        folder_display = Path(main_folder_path).name or str(main_folder_path)
        favicon_uri = get_logo_favicon_b64()
        favicon_tag = report_head_tags(favicon_uri, APP_NAME)

        th_html = "".join(
            f'<th class="col-{c["id"]} align-{c["align"]}">{html_module.escape(c["name"])}</th>'
            for c in active_cols
        )

        sections_html = []
        for folder in folder_data:
            files_in_folder = folder["files"]
            folder_label = folder["folder"]
            folder_size_str = format_inventory_size(folder["total_size"], size_unit)
            file_count = len(files_in_folder)

            rows_html = []
            if not files_in_folder:
                rows_html.append(
                    f'<tr><td colspan="{len(active_cols)}" class="empty-folder-cell">'
                    'No matching files in this folder</td></tr>'
                )
            else:
                for f_info in files_in_folder:
                    cells_html = []
                    for c in active_cols:
                        cid = c["id"]
                        val = f_info.get(cid, "—")
                        if cid == "size" and "size_bytes" in f_info:
                            val = format_inventory_size(f_info["size_bytes"], size_unit)
                        align_cls = f"align-{c['align']}"

                        if cid == "si":
                            cells_html.append(f'<td class="cell-si {align_cls}">{html_module.escape(str(val))}</td>')
                        elif cid == "format":
                            cells_html.append(
                                f'<td class="cell-format {align_cls}">'
                                f'<span class="format-pill">{html_module.escape(str(val))}</span></td>'
                            )
                        elif cid in ("md5", "sha1", "sha256"):
                            if str(val) != "—" and val:
                                cells_html.append(
                                    f'<td class="cell-hash {align_cls}">'
                                    f'<span class="hash-pill" data-hash="{html_module.escape(str(val))}" onclick="copyHash(this, this.dataset.hash)" '
                                    f'title="Click to copy hash">{html_module.escape(str(val))}</span></td>'
                                )
                            else:
                                cells_html.append(f'<td class="cell-hash {align_cls}">—</td>')
                        elif cid == "name":
                            cells_html.append(
                                f'<td class="cell-name {align_cls}" title="{html_module.escape(f_info.get("path", ""))}">'
                                f'<span class="file-icon">{report_icon("file")}</span> '
                                f'<span class="file-name-text">{html_module.escape(str(val))}</span></td>'
                            )
                        else:
                            cells_html.append(f'<td class="cell-{cid} {align_cls}">{html_module.escape(str(val))}</td>')

                    rows_html.append(f'<tr class="data-row">{"".join(cells_html)}</tr>')

            tbody_content = "\n".join(rows_html)
            sec_html = f"""
            <details class="folder-group" open data-folder="{html_module.escape(folder_label)}">
                <summary class="folder-summary">
                    <div class="folder-title">
                        <span class="folder-icon">{report_icon("folder")}</span>
                        <span class="folder-name">{html_module.escape(folder_label)}</span>
                    </div>
                    <div class="folder-stats">
                        <span class="badge file-badge" data-count="{file_count}">{file_count} files</span>
                        <span class="badge size-badge">{folder_size_str}</span>
                    </div>
                </summary>
                <div class="table-container">
                    <table class="inventory-table">
                        <thead>
                            <tr>{th_html}</tr>
                        </thead>
                        <tbody>
                            {tbody_content}
                        </tbody>
                    </table>
                </div>
            </details>
            """
            sections_html.append(sec_html)

        all_sections_str = "\n".join(sections_html)

        html_document = f"""<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{APP_NAME} — File Inventory Report — {html_module.escape(folder_display)}</title>
    {favicon_tag}
    <style>
        {REPORT_BRAND_CSS}
        :root {{
            --bg-color: #0b0e17;
            --card-bg: #161b2b;
            --card-border: #232b43;
            --text-color: #f1f5f9;
            --text-muted: #8e98b0;
            --accent: #3b82f6;
            --accent-hover: #60a5fa;
            --header-bg: #111422;
            --folder-bg: #1e293b;
            --table-header-bg: #232b43;
            --stripe-odd: #111422;
            --stripe-even: #161b2b;
            --hover-row: #1e263d;
            --si-bg: #1e253b;
            --border-subtle: #232b43;
            --pill-bg: #20283f;
            --pill-text: #93c5fd;
            --hash-bg: #0e121d;
            --hash-text: #60a5fa;
            --input-bg: #0e121d;
            --input-border: #232b43;
        }}
        [data-theme="light"] {{
            --bg-color: #f8fafc;
            --card-bg: #ffffff;
            --card-border: #e2e8f0;
            --text-color: #0f172a;
            --text-muted: #64748b;
            --accent: #2563eb;
            --accent-hover: #1d4ed8;
            --header-bg: #ffffff;
            --folder-bg: #3b82f6;
            --table-header-bg: #1e293b;
            --stripe-odd: #ffffff;
            --stripe-even: #f8fafc;
            --hover-row: #eff6ff;
            --si-bg: #dbeafe;
            --border-subtle: #e2e8f0;
            --pill-bg: #dbeafe;
            --pill-text: #1e40af;
            --hash-bg: #f1f5f9;
            --hash-text: #2563eb;
            --input-bg: #ffffff;
            --input-border: #cbd5e1;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-color);
            line-height: 1.5;
            padding: 24px;
            min-height: 100vh;
            transition: background-color 0.2s, color 0.2s;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .report-header {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
        }}
        .header-top {{ display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 16px; margin-bottom: 16px; }}
        .brand-title {{ display: flex; align-items: center; gap: 12px; }}
        .brand-title h1 {{ font-size: 1.6rem; font-weight: 700; color: var(--text-color); }}
        .version-badge {{ background: var(--accent); color: #ffffff; font-size: 0.75rem; font-weight: 700; padding: 3px 10px; border-radius: 9999px; }}
        .theme-toggle-btn {{ background: var(--input-bg); border: 1px solid var(--input-border); color: var(--text-color); padding: 8px 14px; border-radius: 8px; font-size: 0.85rem; cursor: pointer; font-weight: 600; }}
        .meta-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; margin-top: 16px; }}
        .meta-card {{ background: var(--header-bg); border: 1px solid var(--border-subtle); border-radius: 8px; padding: 12px 16px; }}
        .meta-card .label {{ font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--text-muted); font-weight: 600; margin-bottom: 4px; }}
        .meta-card .value {{ font-size: 1.1rem; font-weight: 700; color: var(--text-color); word-break: break-all; }}
        .toolbar {{ background: var(--card-bg); border: 1px solid var(--card-border); border-radius: 12px; padding: 14px 20px; margin-bottom: 24px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 14px; }}
        .search-box {{ display: flex; align-items: center; gap: 10px; flex: 1; min-width: 260px; max-width: 500px; }}
        .search-input {{ width: 100%; background: var(--input-bg); border: 1px solid var(--input-border); color: var(--text-color); padding: 10px 14px; border-radius: 8px; font-size: 0.9rem; outline: none; }}
        .search-input:focus {{ border-color: var(--accent); box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.25); }}
        .match-badge {{ font-size: 0.8rem; font-weight: 600; color: var(--accent); display: none; white-space: nowrap; }}
        .toolbar-buttons {{ display: flex; align-items: center; gap: 10px; }}
        .btn {{ display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px; border-radius: 8px; font-size: 0.85rem; font-weight: 600; cursor: pointer; border: none; }}
        .btn-primary {{ background: var(--accent); color: #ffffff; }}
        .btn-secondary {{ background: var(--input-bg); border: 1px solid var(--input-border); color: var(--text-color); }}
        .folder-group {{ background: var(--card-bg); border: 1px solid var(--card-border); border-radius: 10px; margin-bottom: 18px; overflow: hidden; }}
        .folder-summary {{ background: var(--folder-bg); color: #ffffff; padding: 12px 18px; font-size: 0.95rem; font-weight: 700; cursor: pointer; display: flex; align-items: center; justify-content: space-between; user-select: none; }}
        .folder-title {{ display: flex; align-items: center; gap: 10px; }}
        .folder-stats {{ display: flex; align-items: center; gap: 8px; }}
        .badge {{ background: rgba(255, 255, 255, 0.15); color: #ffffff; font-size: 0.75rem; font-weight: 600; padding: 2px 10px; border-radius: 9999px; }}
        .table-container {{ overflow-x: auto; max-height: 600px; overflow-y: auto; }}
        .inventory-table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; }}
        .inventory-table thead th {{ position: sticky; top: 0; background: var(--table-header-bg); color: #ffffff; font-weight: 600; padding: 10px 12px; z-index: 2; border-bottom: 1px solid var(--border-subtle); white-space: nowrap; }}
        .inventory-table tbody tr {{ border-bottom: 1px solid var(--border-subtle); }}
        .inventory-table tbody tr:nth-child(odd) {{ background-color: var(--stripe-odd); }}
        .inventory-table tbody tr:nth-child(even) {{ background-color: var(--stripe-even); }}
        .inventory-table tbody tr:hover {{ background-color: var(--hover-row); }}
        .inventory-table td {{ padding: 8px 12px; vertical-align: middle; }}
        .align-center {{ text-align: center; }}
        .align-left {{ text-align: left; }}
        .cell-si {{ background-color: var(--si-bg) !important; font-weight: 700; color: var(--accent) !important; width: 50px; }}
        .format-pill {{ background: var(--pill-bg); color: var(--pill-text); font-size: 0.75rem; font-weight: 700; padding: 2px 8px; border-radius: 4px; display: inline-block; }}
        .hash-pill {{ background: var(--hash-bg); color: var(--hash-text); font-family: ui-monospace, SFMono-Regular, monospace; font-size: 0.78rem; padding: 2px 6px; border-radius: 4px; cursor: pointer; border: 1px solid var(--border-subtle); display: inline-block; word-break: break-all; }}
        .empty-folder-cell {{ text-align: center; font-style: italic; color: var(--text-muted); padding: 16px !important; }}
        #toast {{ position: fixed; bottom: 24px; right: 24px; background: var(--accent); color: #ffffff; padding: 10px 18px; border-radius: 8px; font-weight: 600; font-size: 0.85rem; box-shadow: 0 4px 14px rgba(0, 0, 0, 0.4); display: none; z-index: 999; }}
        @media print {{
            body {{ background-color: #ffffff !important; color: #000000 !important; padding: 0 !important; }}
            .toolbar, .theme-toggle-btn, .brand-title .version-badge, #toast {{ display: none !important; }}
            .report-header, .folder-group {{ border: 1px solid #d1d5db !important; box-shadow: none !important; background: #ffffff !important; }}
            details {{ display: block !important; }}
            .folder-summary {{ background: #2563eb !important; color: #ffffff !important; }}
            .inventory-table thead th {{ background: #1e293b !important; color: #ffffff !important; position: static !important; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header class="report-header">
            <div class="header-top">
                <div class="brand-title">
                    {report_brand_header(favicon_uri, APP_NAME, APP_SUBTITLE, APP_VERSION, folder_display, heading=True)}
                    <span class="version-badge">File Inventory</span>
                </div>
                <button class="theme-toggle-btn" id="themeToggleBtn" onclick="toggleTheme()" aria-label="Switch to dark theme">
                    <span class="theme-light-icon">{report_icon("sun")}</span>
                    <span class="theme-dark-icon">{report_icon("moon")}</span>
                    <span id="themeToggleLabel">Dark Mode</span>
                </button>
            </div>
            <div class="meta-grid">
                <div class="meta-card">
                    <div class="label">Source Folder</div>
                    <div class="value" style="font-size: 0.95rem;">{html_module.escape(str(main_folder_path))}</div>
                </div>
                <div class="meta-card"><div class="label">Total Folders</div><div class="value">{len(folder_data):,}</div></div>
                <div class="meta-card"><div class="label">Total Files</div><div class="value">{total_files:,}</div></div>
                <div class="meta-card"><div class="label">Total Size</div><div class="value">{format_inventory_size(total_size, size_unit)}</div></div>
                <div class="meta-card"><div class="label">Generated On</div><div class="value" style="font-size: 0.95rem;">{gen_time}</div></div>
            </div>
        </header>

        <div class="toolbar">
            <div class="search-box">
                <span>{report_icon("search")}</span>
                <input type="text" id="searchInput" class="search-input" placeholder="Search filenames, formats, paths, hashes..." oninput="handleSearch()">
                <span id="matchBadge" class="match-badge"></span>
            </div>
            <div class="toolbar-buttons">
                <button class="btn btn-secondary" onclick="toggleAllFolders(true)">Expand All</button>
                <button class="btn btn-secondary" onclick="toggleAllFolders(false)">Collapse All</button>
                <button class="btn btn-secondary" onclick="exportToCsv()">{report_icon("download")} Export to CSV</button>
                <button class="btn btn-primary" onclick="window.print()">{report_icon("printer")} Print Report</button>
            </div>
        </div>

        <main id="foldersContainer">
            {all_sections_str}
        </main>
        <footer class="report-footer">{APP_NAME} — {html_module.escape(APP_SUBTITLE)} • v{APP_VERSION} • {APP_TAGLINE}</footer>
    </div>
    <div id="toast">Copied to clipboard!</div>
    <script>
        function handleSearch() {{
            const query = document.getElementById('searchInput').value.toLowerCase().trim();
            const folderGroups = document.querySelectorAll('.folder-group');
            const matchBadge = document.getElementById('matchBadge');
            let totalMatches = 0; let totalRows = 0;
            folderGroups.forEach(group => {{
                const rows = group.querySelectorAll('tbody tr.data-row');
                let count = 0;
                rows.forEach(row => {{
                    totalRows++;
                    const text = row.innerText.toLowerCase();
                    if (!query || text.includes(query)) {{
                        row.style.display = ''; count++; totalMatches++;
                    }} else {{ row.style.display = 'none'; }}
                }});
                const fileBadge = group.querySelector('.file-badge');
                if (query) {{
                    if (count > 0) {{ group.style.display = ''; group.open = true; if (fileBadge) fileBadge.innerText = count + ' matching'; }}
                    else {{ group.style.display = 'none'; }}
                }} else {{
                    group.style.display = '';
                    if (fileBadge) fileBadge.innerText = fileBadge.getAttribute('data-count') + ' files';
                }}
            }});
            if (query) {{
                matchBadge.style.display = 'inline-block';
                matchBadge.innerText = 'Showing ' + totalMatches + ' of ' + totalRows + ' files';
            }} else {{ matchBadge.style.display = 'none'; }}
        }}
        function toggleAllFolders(expand) {{
            document.querySelectorAll('.folder-group').forEach(group => group.open = expand);
        }}
        function toggleTheme() {{
            const root = document.documentElement;
            const cur = root.getAttribute('data-theme') || 'dark';
            const next = cur === 'dark' ? 'light' : 'dark';
            root.setAttribute('data-theme', next);
            document.getElementById('themeToggleLabel').textContent = next === 'dark' ? 'Light Mode' : 'Dark Mode';
            document.getElementById('themeToggleBtn').setAttribute('aria-label', 'Switch to ' + (next === 'dark' ? 'light' : 'dark') + ' theme');
        }}
        function copyHash(element, text) {{
            navigator.clipboard.writeText(text).then(() => {{
                element.style.background = '#10b981'; element.style.color = '#ffffff';
                showToast();
                setTimeout(() => {{ element.style.background = ''; element.style.color = ''; }}, 1200);
            }});
        }}
        function showToast() {{
            const t = document.getElementById('toast'); t.style.display = 'block';
            setTimeout(() => t.style.display = 'none', 1800);
        }}
        function exportToCsv() {{
            const rows = []; const headers = ['Folder'];
            const firstHeaderRow = document.querySelector('.inventory-table thead tr');
            if (firstHeaderRow) {{
                firstHeaderRow.querySelectorAll('th').forEach(th => headers.push(th.innerText.trim()));
            }}
            rows.push(headers.map(escapeCsvValue).join(','));
            document.querySelectorAll('.folder-group').forEach(group => {{
                const folderName = group.getAttribute('data-folder') || '';
                group.querySelectorAll('tbody tr.data-row').forEach(row => {{
                    const rowData = [folderName];
                    row.querySelectorAll('td').forEach(td => rowData.push(td.innerText.trim()));
                    rows.push(rowData.map(escapeCsvValue).join(','));
                }});
            }});
            const blob = new Blob([rows.join('\\r\\n')], {{ type: 'text/csv;charset=utf-8;' }});
            const link = document.createElement('a'); link.href = URL.createObjectURL(blob);
            link.download = {script_string(folder_display + '_File_Inventory.csv')};
            document.body.appendChild(link); link.click(); document.body.removeChild(link);
        }}
        function escapeCsvValue(val) {{
            if (val === null || val === undefined) return '""';
            if (/^[\\s]*[=+@-]/.test(String(val))) val = "'" + val;
            return '"' + String(val).replace(/"/g, '""') + '"';
        }}
    </script>
</body>
</html>"""

        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        with new_report(output_file) as f:
            f.write(html_document)


class HashReportExporter:
    """Exports hash verification results into Checksum files, Excel, CSV, or HTML."""

    @staticmethod
    def export_checksum_file(output_path, file_items, default_algorithm="SHA-256"):
        """Use paths relative to the manifest so it can be imported again."""
        output = Path(output_path).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        lines = []
        for item in file_items:
            if item.get("status") not in ("match", "mismatch", "computed"):
                continue
            digest = item.get("calculated_hash")
            if not digest:
                continue
            source = Path(item["path"]).resolve()
            try:
                filename = os.path.relpath(source, output.parent)
            except ValueError:  # A different Windows drive needs an absolute path.
                filename = str(source)
            if "\n" in filename or "\r" in filename:
                raise ValueError("Checksum filenames must not contain newlines")
            lines.append(f"{digest}  {filename}")
        write_report(output, "\n".join(lines) + ("\n" if lines else ""))

    @staticmethod
    def export_excel(output_path, file_items, default_algorithm="SHA-256"):
        if not OPENPYXL_AVAILABLE:
            raise _openpyxl_unavailable_error(
                "openpyxl is required to export Excel reports."
            ) from OPENPYXL_IMPORT_ERROR

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Hash Verification"
        ws.sheet_view.showGridLines = False

        headers = [
            "SI No.", "Status", "File Name", "File Size",
            "Algorithm", "Expected Hash", "Calculated Hash", "Result", "Full Path"
        ]
        widths = [8, 18, 40, 14, 12, 40, 40, 14, 60]

        # Row 1: Title
        ws.merge_cells("A1:I1")
        ws["A1"] = "HASH VERIFICATION & INTEGRITY REPORT"
        ws["A1"].font = XlFont(name="Segoe UI", size=14, bold=True, color=XL_COLOR_WHITE)
        ws["A1"].fill = PatternFill("solid", fgColor=XL_COLOR_TITLE_BG)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Row 2: Metadata
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        matches = sum(1 for it in file_items if it.get("result") == "MATCH")
        mismatches = sum(1 for it in file_items if it.get("result") == "MISMATCH")
        meta_str = f"Generated: {now_str}    |    Total: {len(file_items)}    |    Matches: {matches}    |    Mismatches: {mismatches}"
        ws.merge_cells("A2:I2")
        ws["A2"] = meta_str
        ws["A2"].font = XlFont(name="Segoe UI", size=9, bold=True, color=XL_COLOR_HEADER_BG)
        ws["A2"].fill = PatternFill("solid", fgColor=XL_COLOR_SI_BG)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22

        # Row 3: Blank
        ws.row_dimensions[3].height = 10

        # Row 4: Column Headers
        ws.row_dimensions[4].height = 24
        thin_side = Side(style="thin", color=XL_COLOR_BORDER)
        border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=4, column=col_idx, value=h)
            c.font = XlFont(name="Segoe UI", size=10, bold=True, color=XL_COLOR_WHITE)
            c.fill = PatternFill("solid", fgColor=XL_COLOR_HEADER_BG)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border_all

        # Row 5+: Data rows
        for r_idx, item in enumerate(file_items, start=5):
            ws.row_dimensions[r_idx].height = 20
            res = item.get("result", "PENDING")
            status_text = item.get("status_text", "Pending")

            calc_h = item.get("calculated_hash") or item.get("all_hashes", {}).get(item.get('algorithm', default_algorithm), "—")
            exp_h = item.get("expected_hash") or "—"

            vals = [
                r_idx - 4,
                status_text,
                item["name"],
                format_size(item.get("size", 0)),
                item.get("algorithm", default_algorithm),
                exp_h,
                calc_h,
                res,
                item["path"]
            ]

            is_even = (r_idx % 2 == 1)
            row_bg = XL_COLOR_STRIPE_EVEN if is_even else XL_COLOR_STRIPE_ODD

            for col_idx, val in enumerate(vals, 1):
                c = ws.cell(row=r_idx, column=col_idx)
                set_excel_literal(c, val)
                c.font = XlFont(name="Segoe UI", size=9)
                c.border = border_all
                c.fill = PatternFill("solid", fgColor=row_bg)

                if col_idx in (1, 4, 5):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in (2, 8):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    if res == "MATCH":
                        c.fill = PatternFill("solid", fgColor=XL_COLOR_SUCCESS_BG)
                        c.font = XlFont(name="Segoe UI", size=9, bold=True, color=XL_COLOR_SUCCESS_TXT)
                    elif res == "MISMATCH":
                        c.fill = PatternFill("solid", fgColor=XL_COLOR_DANGER_BG)
                        c.font = XlFont(name="Segoe UI", size=9, bold=True, color=XL_COLOR_DANGER_TXT)
                    elif res == "NO TARGET":
                        c.fill = PatternFill("solid", fgColor=XL_COLOR_INFO_BG)
                        c.font = XlFont(name="Segoe UI", size=9, color=XL_COLOR_INFO_TXT)
                elif col_idx in (6, 7):
                    c.font = XlFont(name="Consolas", size=8)
                    c.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        for idx, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = w

        ws.freeze_panes = "A5"
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        save_workbook(wb, output_path)

    @staticmethod
    def export_csv(output_path, file_items, default_algorithm="SHA-256"):
        headers = ["SI", "Status", "FileName", "FileSize", "Algorithm", "ExpectedHash", "CalculatedHash", "Result", "FullPath"]
        rows = [headers]
        for idx, item in enumerate(file_items, 1):
            calc_h = item.get("calculated_hash") or item.get("all_hashes", {}).get(item.get('algorithm', default_algorithm), "")
            row = [
                str(idx),
                item.get("status_text", ""),
                item["name"],
                str(item.get("size", 0)),
                item.get("algorithm", default_algorithm),
                item.get("expected_hash", ""),
                calc_h,
                item.get("result", ""),
                item["path"]
            ]
            rows.append(row)

        write_report(output_path, csv_text(rows))

    @staticmethod
    def export_html(output_path, file_items, default_algorithm="SHA-256"):
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        matches = sum(1 for it in file_items if it.get("result") == "MATCH")
        mismatches = sum(1 for it in file_items if it.get("result") == "MISMATCH")
        computed = sum(1 for it in file_items if it.get("result") == "NO TARGET")

        rows_html = []
        for idx, item in enumerate(file_items, 1):
            res = item.get("result", "PENDING")
            if res == "MATCH":
                status_style = "color: #059669; font-weight: 600;"
                status_text = "Verified Match"
            elif res == "MISMATCH":
                status_style = "color: #dc2626; font-weight: 600;"
                status_text = "Mismatch"
            elif res == "ERROR":
                status_style = "color: #dc2626; font-weight: 600;"
                status_text = "Error"
            else:
                status_style = "color: #2563eb; font-weight: 500;"
                status_text = item.get("status_text", "Calculated")

            calc_h = item.get("calculated_hash") or item.get("all_hashes", {}).get(item.get('algorithm', default_algorithm), "—")
            exp_h = item.get("expected_hash") or "—"

            rows_html.append(f"""
            <tr>
                <td style="text-align: center; color: #64748b;">{idx}</td>
                <td style="text-align: center; {status_style}">{html_module.escape(status_text)}</td>
                <td><strong>{html_module.escape(item['name'])}</strong></td>
                <td style="text-align: right; color: #475569;">{format_size(item.get('size', 0))}</td>
                <td style="text-align: center; font-weight: 600;">{html_module.escape(item.get('algorithm', default_algorithm))}</td>
                <td><code class="hash-code">{html_module.escape(exp_h)}</code></td>
                <td><code class="hash-code">{html_module.escape(calc_h)}</code></td>
                <td style="font-size: 11px; color: #64748b;">{html_module.escape(item['path'])}</td>
            </tr>
            """)

        html_body = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{APP_NAME} — Cryptographic Hash Verification Report</title>
    {report_head_tags(get_logo_favicon_b64(), APP_NAME)}
    <style>
        {REPORT_BRAND_CSS}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background: #ffffff;
            color: #0f172a;
            padding: 36px 48px;
            font-size: 13px;
            line-height: 1.5;
        }}
        .report-header {{
            border-bottom: 2px solid #0f172a;
            padding-bottom: 12px;
            margin-bottom: 16px;
        }}
        h1 {{
            font-size: 18px;
            font-weight: 700;
            margin: 0 0 6px 0;
            color: #0f172a;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }}
        .header-meta {{
            font-size: 12px;
            color: #64748b;
        }}
        .summary-stats {{
            margin-bottom: 20px;
            font-size: 13px;
            color: #334155;
            padding: 8px 14px;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            margin-top: 8px;
        }}
        th {{
            background: #f8fafc;
            color: #334155;
            font-weight: 700;
            padding: 8px 10px;
            text-align: left;
            border-top: 1px solid #cbd5e1;
            border-bottom: 1px solid #cbd5e1;
        }}
        td {{
            padding: 7px 10px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
            color: #0f172a;
        }}
        tr:hover {{ background: #f8fafc; }}
        .hash-code {{
            font-family: Consolas, "Courier New", monospace;
            font-size: 11px;
            word-break: break-all;
            color: #0f172a;
        }}
        .report-footer {{
            margin-top: 40px;
            padding-top: 12px;
            border-top: 1px solid #e2e8f0;
            font-size: 11px;
            color: #94a3b8;
            text-align: center;
        }}
        @media print {{
            body {{ padding: 15px; font-size: 11px; }}
            h1 {{ font-size: 15px; }}
            th {{ background: #f0f0f0 !important; }}
        }}
    </style>
</head>
<body>
    <div class="report-header">
        {report_brand_header(get_logo_favicon_b64(), APP_NAME, APP_SUBTITLE, APP_VERSION)}
        <h1>Cryptographic Hash Verification Report</h1>
        <div class="header-meta">Generated: {now_str} • Primary Algorithm: {default_algorithm}</div>
    </div>
    <div class="summary-stats">
        <strong>Total Files:</strong> {len(file_items)} &nbsp;|&nbsp;
        <strong>Verified Matches:</strong> <span style="color: #059669; font-weight: 600;">{matches}</span> &nbsp;|&nbsp;
        <strong>Mismatches:</strong> <span style="color: #dc2626; font-weight: 600;">{mismatches}</span> &nbsp;|&nbsp;
        <strong>Calculated (Unchecked):</strong> <span style="color: #2563eb; font-weight: 600;">{computed}</span>
    </div>
    <table>
        <thead>
            <tr>
                <th style="width: 40px; text-align: center;">#</th>
                <th style="width: 120px; text-align: center;">Status</th>
                <th style="width: 220px;">File Name</th>
                <th style="width: 90px; text-align: right;">Size</th>
                <th style="width: 80px; text-align: center;">Algorithm</th>
                <th>Expected Hash</th>
                <th>Calculated Hash</th>
                <th>Full Path</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows_html)}
        </tbody>
    </table>
    <div class="report-footer">
        {APP_NAME} — {html_module.escape(APP_SUBTITLE)} • v{APP_VERSION} • {APP_TAGLINE}
    </div>
</body>
</html>"""
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with new_report(output_path) as f:
            f.write(html_body)


class MetadataReportExporter:
    """
    Generates professional forensic Metadata Reports in HTML, TXT, and JSON formats
    with strict attribution identifying ExifTool and its dynamic version.
    """

    ATTRIBUTION_TEXT = (
        "Metadata extraction was performed using ExifTool. The metadata presented in this "
        "report is derived from the output generated by ExifTool version {exiftool_version}. "
        "Hash values, where included, were calculated independently using the application's "
        "existing hash calculation engine."
    )

    @classmethod
    def export_html(cls, output_file, record, integrity_info=None, include_hashes=True):
        """Generates minimal, standard, professional white-theme HTML metadata extraction report."""
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file_name = record.get("file_name", Path(record.get("file_path", "")).name)
        file_path = record.get("file_path", "")
        file_size_bytes = record.get("file_size", 0)
        file_size_str = f"{file_size_bytes:,} bytes ({file_size_bytes / (1024*1024):.2f} MB)" if file_size_bytes else format_size(file_size_bytes)
        exiftool_ver = record.get("exiftool_version", "13.59")

        # Cryptographic Hash Verification Table (clean, minimal)
        hashes_html = ""
        if include_hashes and integrity_info:
            pre_h = integrity_info.get("pre") or {}
            post_h = integrity_info.get("post") or {}
            verified = integrity_info.get("verified", False)
            status_str = integrity_status_text(integrity_info)
            status_color = "#059669" if verified is True else ("#64748b" if verified is None else "#dc2626")

            rows = []
            for algo in ["SHA-256", "SHA-512", "MD5", "SHA-1"]:
                if algo not in pre_h and algo not in post_h:
                    continue
                val_pre = pre_h.get(algo, "—")
                val_post = post_h.get(algo, "—")
                if algo in pre_h and algo in post_h:
                    matched = pre_h[algo] == post_h[algo]
                    status_str = "Verified Match" if matched else "Mismatch Detected"
                    status_color = "#059669" if matched else "#dc2626"
                else:
                    status_str, status_color = "Not compared", "#64748b"
                rows.append(f"""
                <tr>
                    <td style="font-weight: 600; width: 120px;">{algo}</td>
                    <td><code class="hash-code">{html_module.escape(val_pre)}</code></td>
                    <td><code class="hash-code">{html_module.escape(val_post)}</code></td>
                    <td style="text-align: center; color: {status_color}; font-weight: 600;">{status_str}</td>
                </tr>
                """)

            hashes_html = f"""
            <h2>Cryptographic Hash Verification</h2>
            <table class="meta-table">
                <thead>
                    <tr>
                        <th style="width: 120px;">Algorithm</th>
                        <th>Pre-Extraction Checksum</th>
                        <th>Post-Extraction Checksum</th>
                        <th style="width: 150px; text-align: center;">Integrity Status</th>
                    </tr>
                </thead>
                <tbody>
                    {"".join(rows)}
                </tbody>
            </table>
            """

        # Metadata Categories Tables (minimal, clean standard)
        categories = record.get("categories", {})
        cat_cards = []
        for cat_name, items in categories.items():
            if not items:
                continue
            rows = []
            for it in items:
                val = it.get("value")
                row_class = 'structured-row' if isinstance(val, (dict, list)) or len(str(val)) > 600 else ''
                if isinstance(val, (dict, list)):
                    val_str = json.dumps(val, ensure_ascii=False, indent=2)
                    lines = ''.join(f"<div class='struct-line'>{html_module.escape(line)}</div>" for line in val_str.splitlines())
                    val_html = f"<div class='struct-pre'>{lines}</div>"
                else:
                    val_html = html_module.escape(str(val))
                rows.append(f"""
                <tr class="{row_class}">
                    <td class="tag-name">{html_module.escape(it.get('tag', ''))}</td>
                    <td class="tag-val">{val_html}</td>
                    <td class="tag-key">{html_module.escape(it.get('raw_key', ''))}</td>
                </tr>
                """)

            cat_cards.append(f"""
            <h2>{html_module.escape(cat_name)} ({len(items)} tags)</h2>
            <table class="meta-table">
                <thead>
                    <tr>
                        <th style="width: 24%;">Tag Name</th>
                        <th>Value</th>
                        <th style="width: 26%;">ExifTool Key</th>
                    </tr>
                </thead>
                <tbody>
                    {"".join(rows)}
                </tbody>
            </table>
            """)

        # Extraction Warnings (if any)
        warnings = record.get("warnings", [])
        warn_html = ""
        if warnings:
            w_items = "".join(f"<li>{html_module.escape(w)}</li>" for w in warnings)
            warn_html = f"""
            <h2>Extraction Warnings</h2>
            <ul style="font-size: 12px; color: #b45309; padding-left: 20px;">
                {w_items}
            </ul>
            """

        html_body = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{APP_NAME} — Metadata Report — {html_module.escape(file_name)}</title>
    {report_head_tags(get_logo_favicon_b64(), APP_NAME)}
    <style>
        {REPORT_BRAND_CSS}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #ffffff;
            color: #0f172a;
            margin: 0;
            padding: 36px 48px;
            font-size: 13px;
            line-height: 1.5;
        }}
        .report-header {{
            border-bottom: 2px solid #0f172a;
            padding-bottom: 12px;
            margin-bottom: 20px;
        }}
        h1 {{
            font-size: 18px;
            font-weight: 700;
            margin: 0 0 6px 0;
            color: #0f172a;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }}
        .header-meta {{
            font-size: 12px;
            color: #64748b;
        }}
        .file-summary-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}
        .file-summary-table td {{
            padding: 5px 10px;
            font-size: 13px;
            vertical-align: top;
            border-bottom: 1px solid #f1f5f9;
        }}
        .file-summary-table td.lbl {{
            width: 150px;
            font-weight: 600;
            color: #475569;
        }}
        .file-summary-table td.val {{
            color: #0f172a;
            word-break: break-all;
        }}
        h2 {{
            font-size: 13px;
            font-weight: 700;
            color: #0f172a;
            margin: 18px 0 6px 0;
            padding-bottom: 5px;
            border-bottom: 1px solid #cbd5e1;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        table.meta-table {{
            width: 100%;
            table-layout: fixed;
            border-collapse: collapse;
            margin-bottom: 14px;
            font-size: 12px;
        }}
        table.meta-table th {{
            background-color: #f8fafc;
            color: #334155;
            font-weight: 700;
            text-align: left;
            padding: 6px 8px;
            border-top: 1px solid #cbd5e1;
            border-bottom: 1px solid #cbd5e1;
        }}
        table.meta-table td {{
            padding: 5px 8px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
            color: #0f172a;
        }}
        table.meta-table tr:hover {{
            background-color: #f8fafc;
        }}
        .tag-name {{
            font-weight: 600;
            color: #0f172a;
            width: 24%;
        }}
        .tag-val {{
            overflow-wrap: anywhere;
            word-break: normal;
            color: #0f172a;
        }}
        .tag-key {{
            font-family: Consolas, "Courier New", monospace;
            font-size: 11px;
            color: #64748b;
            width: 26%;
            overflow-wrap: anywhere;
        }}
        .hash-code {{
            font-family: Consolas, "Courier New", monospace;
            font-size: 11px;
            color: #0f172a;
            word-break: break-all;
        }}
        .struct-pre {{
            font-family: Consolas, "Courier New", monospace;
            font-size: 11px;
            margin: 0;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 4px;
            padding: 6px 8px;
            max-height: 180px;
            overflow-y: auto;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        .report-footer {{
            margin-top: 40px;
            padding-top: 12px;
            border-top: 1px solid #e2e8f0;
            font-size: 11px;
            color: #94a3b8;
            text-align: center;
        }}
        @page {{ size: A4 portrait; margin: 14mm 12mm 15mm; }}
        @media print {{
            html, body {{ width: auto; margin: 0; padding: 0; }}
            body {{ font-size: 9pt; line-height: 1.3; }}
            .report-header {{ padding-bottom: 6px; margin-bottom: 10px; }}
            h1 {{ font-size: 13pt; }}
            h2 {{ font-size: 10pt; margin: 12px 0 5px; break-after: avoid; page-break-after: avoid; }}
            .file-summary-table {{ margin-bottom: 10px; table-layout: fixed; }}
            .file-summary-table td {{ padding: 3px 6px; font-size: 9pt; overflow-wrap: anywhere; }}
            .file-summary-table td.lbl {{ width: 24%; }}
            table.meta-table {{ font-size: 8pt; margin-bottom: 10px; table-layout: fixed; }}
            table.meta-table th, table.meta-table td {{ padding: 4px 6px; overflow-wrap: anywhere; }}
            table.meta-table th {{ background: #f0f0f0 !important; }}
            thead {{ display: table-header-group; }}
            tr {{ break-inside: avoid; page-break-inside: avoid; }}
            tr.structured-row, .struct-pre {{ break-inside: auto; page-break-inside: auto; }}
            .struct-line {{ break-inside: avoid; page-break-inside: avoid; }}
            .hash-code, .tag-key, .struct-pre {{ font-size: 7.5pt; overflow-wrap: anywhere; }}
            .struct-pre {{ max-height: none; overflow: visible; white-space: pre-wrap; }}
            .report-footer {{ margin-top: 8px; padding-top: 6px; font-size: 8pt; }}
        }}
    </style>
</head>
<body>
    <div class="report-header">
        {report_brand_header(get_logo_favicon_b64(), APP_NAME, APP_SUBTITLE, APP_VERSION)}
        <h1>Metadata Extraction Report</h1>
        <div class="header-meta">Generated: {now_str} • Extraction Engine: ExifTool v{exiftool_ver}</div>
    </div>

    <table class="file-summary-table">
        <tr>
            <td class="lbl">Source File</td>
            <td class="val"><strong>{html_module.escape(file_name)}</strong></td>
        </tr>
        <tr>
            <td class="lbl">Full Path</td>
            <td class="val">{html_module.escape(file_path)}</td>
        </tr>
        <tr>
            <td class="lbl">File Size</td>
            <td class="val">{file_size_str}</td>
        </tr>
    </table>

    {hashes_html}
    {warn_html}
    {"".join(cat_cards)}

    <div class="report-footer">
        {APP_NAME} — {html_module.escape(APP_SUBTITLE)} • v{APP_VERSION} • {APP_TAGLINE}
    </div>
</body>
</html>"""
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        with new_report(output_file) as f:
            f.write(html_body)
        return output_file

    @classmethod
    def export_txt(cls, output_file, record, integrity_info=None, include_hashes=True):
        """Generates structured human-readable plain-text forensic report."""
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file_name = record.get("file_name", Path(record.get("file_path", "")).name)
        file_path = record.get("file_path", "")
        file_size = format_size(record.get("file_size", 0))
        exiftool_version = record.get("exiftool_version", "Unknown")

        lines = [
            "=" * 80,
            "METADATA EXTRACTION REPORT",
            f"{APP_NAME} {APP_VERSION} — Forensic Metadata Examination",
            "=" * 80,
            f"Report Generated : {now_str}",
            f"Extraction Engine: ExifTool v{exiftool_version}",
            f"Source File Name : {file_name}",
            f"Source File Path : {file_path}",
            f"File Size        : {file_size}",
            "-" * 80,
        ]

        if include_hashes and integrity_info:
            pre_h = integrity_info.get("pre") or {}
            post_h = integrity_info.get("post") or {}
            verified = integrity_info.get("verified", False)
            status_text = integrity_status_text(integrity_info).upper()

            lines.extend([
                "CRYPTOGRAPHIC HASH INTEGRITY VERIFICATION:",
                f"Status: {status_text}",
                f"SHA-256 (Pre)  : {pre_h.get('SHA-256', '—')}",
                f"SHA-256 (Post) : {post_h.get('SHA-256', '—')}",
                f"SHA-512 (Pre)  : {pre_h.get('SHA-512', '—')}",
                f"MD5 (Pre)      : {pre_h.get('MD5', '—')}",
                f"SHA-1 (Pre)    : {pre_h.get('SHA-1', '—')}",
                "-" * 80,
            ])
        elif not include_hashes:
            lines.extend([
                "CRYPTOGRAPHIC HASHING: Not requested by user.",
                "-" * 80,
            ])

        warnings = record.get("warnings", [])
        if warnings:
            lines.append("EXTRACTION WARNINGS / ENGINE MESSAGES:")
            for w in warnings:
                lines.append(f"  [!] {w}")
            lines.append("-" * 80)

        categories = record.get("categories", {})
        for cat_name, items in categories.items():
            if not items:
                continue
            lines.append(f"[{cat_name.upper()}] ({len(items)} tags)")
            for it in items:
                val = it.get("value")
                val_repr = json.dumps(val) if isinstance(val, (dict, list)) else str(val)
                lines.append(f"  {it.get('tag', ''):<32} : {val_repr}")
            lines.append("")

        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        with new_report(output_file) as f:
            f.write("\n".join(lines))
        return output_file

    @classmethod
    def export_json(cls, output_dir, record, integrity_info=None, include_hashes=True, preserve_raw=True):
        """
        Preserves original raw ExifTool JSON output separately as [basename]_exiftool_raw.json
        AND creates application-level structured record as [basename]_metadata_record.json.
        """
        p_dir = Path(output_dir)
        p_dir.mkdir(parents=True, exist_ok=True)
        fname = record.get("file_name", Path(record.get("file_path", "evidence")).name)
        stem = metadata_report_basename(record)
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        raw_path = None
        if preserve_raw:
            raw_path = p_dir / f"{stem}_exiftool_raw.json"
            raw_json_str = record.get("raw_json", "[]")
            with new_report(raw_path) as f:
                f.write(raw_json_str)

        record_path = p_dir / f"{stem}_metadata_record.json"
        app_record = {
            "application": f"{APP_NAME} {APP_VERSION}",
            "generated_at": now_str,
            "extraction_engine": "ExifTool",
            "exiftool_version": record.get("exiftool_version", "Unknown"),
            "source_file": {
                "name": fname,
                "path": record.get("file_path", ""),
                "size_bytes": record.get("file_size", 0),
                "size_formatted": format_size(record.get("file_size", 0)),
            },
            "integrity_verification": {
                "hashes_enabled": include_hashes,
                "verified": integrity_info.get("verified", False) if integrity_info else None,
                "status": integrity_info.get("status", "NOT_REQUESTED") if integrity_info else "NOT_REQUESTED",
                "hashes_pre": integrity_info.get("pre", {}) if integrity_info else {},
                "hashes_post": integrity_info.get("post", {}) if integrity_info else {},
            },
            "warnings": record.get("warnings", []),
            "metadata_categories": record.get("categories", {}),
            "parsed_raw_metadata": record.get("parsed", {})
        }
        with new_report(record_path) as f:
            json.dump(app_record, f, ensure_ascii=False, indent=2)

        return str(record_path), str(raw_path) if raw_path else None


# ==============================================================================
# SECTION 5: QT BACKGROUND WORKER THREADS
# ==============================================================================

class InventoryWorker(QThread):
    """
    QThread worker that crawls directories, extracts media durations, computes
    cryptographic hashes, and writes Excel/HTML reports without freezing the UI.
    """
    scan_started = Signal()
    scan_progress = Signal(int, str)
    process_started = Signal(int)
    process_progress = Signal(int, int, str)
    writing_report = Signal(str)
    inventory_finished = Signal(list, dict, list)
    error_occurred = Signal(str, str)
    inventory_cancelled = Signal(int)
    no_files_found = Signal(list)

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self._cancel_event = threading.Event()

    def cancel(self):
        self._cancel_event.set()

    def run(self):
        try:
            with report_write_scope(approved=self.config.get('approved_replacements'),
                                    protected_paths=self.config.get('protected_paths', ()),
                                    protected_roots=(self.config['source_folder'],)):
                self._run_inventory()
        except Exception as exc:
            self.error_occurred.emit("Inventory Error", str(exc))

    def _run_inventory(self):
        source_dir = Path(self.config["source_folder"])
        if not source_dir.is_dir():
            raise ValueError("Source directory no longer exists")
        out_excel = self.config.get("output_excel")
        out_html = self.config.get("output_html")
        format_type = self.config["format_type"]
        filter_cats = self.config["filter_categories"]
        enabled_cols = self.config["enabled_columns"]
        si_mode = self.config["si_mode"]
        skip_empty_folders = self.config.get("skip_empty_folders", False)
        size_unit = self.config.get("size_unit", "auto")

        need_duration = "duration" in enabled_cols
        need_md5 = "md5" in enabled_cols
        need_sha1 = "sha1" in enabled_cols
        need_sha256 = "sha256" in enabled_cols
        need_any_hash = need_md5 or need_sha1 or need_sha256

        skip_paths = set()
        if out_excel:
            try: skip_paths.add(Path(out_excel).resolve())
            except Exception: pass
        if out_html:
            try: skip_paths.add(Path(out_html).resolve())
            except Exception: pass

        allowed_extensions = set()
        include_other = filter_cats.get("other", False)
        if filter_cats.get("audio"): allowed_extensions.update(AUDIO_EXTENSIONS)
        if filter_cats.get("video"): allowed_extensions.update(VIDEO_EXTENSIONS)
        if filter_cats.get("images"): allowed_extensions.update(IMAGE_EXTENSIONS)
        if filter_cats.get("docs"): allowed_extensions.update(DOCUMENT_EXTENSIONS)
        if filter_cats.get("archives"): allowed_extensions.update(ARCHIVE_EXTENSIONS)

        self.scan_started.emit()

        folder_data = []
        errors = []
        if need_duration and not FFPROBE_BIN:
            errors.append("ffprobe is unavailable; media durations could not be extracted.")
        total_discovered_files = 0
        total_size_bytes = 0
        all_file_records = []
        start_time = time.time()

        for current_folder, dirs, files in evidence_walk(
            str(source_dir),
            onerror=lambda err: errors.append(f"Directory access error: {err}")
        ):
            if self._cancel_event.is_set():
                self.inventory_cancelled.emit(total_discovered_files)
                return

            dirs.sort(key=str.casefold)
            files.sort(key=str.casefold)
            cur_path = Path(current_folder)
            try:
                rel_path = os.path.relpath(current_folder, str(source_dir))
                folder_label = "MAIN FOLDER" if rel_path == "." else rel_path
            except Exception:
                folder_label = current_folder

            folder_files = []
            folder_size = 0

            for fname in files:
                if self._cancel_event.is_set():
                    self.inventory_cancelled.emit(total_discovered_files)
                    return

                if fname.startswith("~$"):
                    continue

                fpath = cur_path / fname
                try:
                    if fpath.is_symlink():
                        continue
                    try:
                        if fpath.resolve() in skip_paths:
                            continue
                    except Exception:
                        pass

                    if not fpath.is_file():
                        continue

                    ext = fpath.suffix.lower()
                    matched = False
                    if ext in allowed_extensions:
                        matched = True
                    elif include_other and (ext not in ALL_NAMED_EXTENSIONS):
                        matched = True

                    if not matched:
                        continue

                    stat = fpath.stat()
                    fsize = stat.st_size
                    folder_size += fsize
                    total_size_bytes += fsize

                    fmt_display = ext[1:].upper() if ext else "NO EXTENSION"
                    mod_dt = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    cre_dt = datetime.datetime.fromtimestamp(stat.st_ctime).strftime("%Y-%m-%d %H:%M:%S")

                    record = {
                        "name": fname,
                        "size": format_inventory_size(fsize, size_unit),
                        "size_bytes": fsize,
                        "format": fmt_display,
                        "duration": "—",
                        "modified": mod_dt,
                        "created": cre_dt,
                        "path": str(fpath.resolve()),
                        "ext": ext,
                        "file_path_obj": fpath,
                        "md5": "—",
                        "sha1": "—",
                        "sha256": "—",
                    }
                    folder_files.append(record)
                    all_file_records.append(record)
                    total_discovered_files += 1

                except (OSError, PermissionError) as ex:
                    errors.append(f"{fpath}: {ex}")

            if skip_empty_folders and not folder_files:
                self.scan_progress.emit(total_discovered_files, folder_label)
                continue

            folder_data.append({
                "folder": folder_label,
                "path": str(cur_path),
                "files": folder_files,
                "total_size": folder_size
            })
            self.scan_progress.emit(total_discovered_files, folder_label)

        if total_discovered_files == 0:
            self.no_files_found.emit(errors)
            return

        need_processing_pass = need_duration or need_any_hash
        processed_count = 0

        if need_processing_pass:
            self.process_started.emit(total_discovered_files)

            for rec in all_file_records:
                if self._cancel_event.is_set():
                    self.inventory_cancelled.emit(processed_count)
                    return

                processed_count += 1
                fpath = rec["file_path_obj"]
                ext = rec["ext"]

                if need_duration and (ext in MEDIA_EXTENSIONS):
                    sec = extract_media_duration(fpath)
                    if sec is not None:
                        rec["duration"] = format_duration(sec)

                if need_any_hash:
                    try:
                        hashes = compute_all_file_hashes(fpath, cancel_event=self._cancel_event, algorithms={name for column, name in (("md5", "MD5"), ("sha1", "SHA-1"), ("sha256", "SHA-256")) if column in enabled_cols})
                        if hashes is None:
                            self.inventory_cancelled.emit(processed_count)
                            return
                        rec["md5"] = hashes.get("MD5", "—")
                        rec["sha1"] = hashes.get("SHA-1", "—")
                        rec["sha256"] = hashes.get("SHA-256", "—")
                    except (OSError, PermissionError) as ex:
                        errors.append(f"Hash calculation failed for {fpath}: {ex}")

                self.process_progress.emit(processed_count, total_discovered_files, rec["name"])

        # Assign Serial Numbers
        continuous_si = 1
        for folder in folder_data:
            folder_si = 1
            for rec in folder["files"]:
                if si_mode == "folder":
                    rec["si"] = folder_si
                    folder_si += 1
                else:
                    rec["si"] = continuous_si
                    continuous_si += 1

        if self._cancel_event.is_set():
            self.inventory_cancelled.emit(processed_count)
            return
        generated_files = []
        try:
            if format_type in ("excel", "both") and out_excel:
                self.writing_report.emit("Excel Workbook (.xlsx)")
                ExcelReportGenerator.generate(
                    output_file=out_excel,
                    folder_data=folder_data,
                    enabled_column_ids=enabled_cols,
                    main_folder_path=str(source_dir),
                    total_files=total_discovered_files,
                    total_size=total_size_bytes,
                    size_unit=size_unit
                )
                generated_files.append(out_excel)

            if format_type in ("html", "both") and out_html:
                self.writing_report.emit("HTML Report (.html)")
                HtmlReportGenerator.generate(
                    output_file=out_html,
                    folder_data=folder_data,
                    enabled_column_ids=enabled_cols,
                    main_folder_path=str(source_dir),
                    total_files=total_discovered_files,
                    total_size=total_size_bytes,
                    size_unit=size_unit
                )
                generated_files.append(out_html)

        except PermissionError as p_err:
            self.error_occurred.emit(
                "Permission Denied",
                f"Cannot save output file: {p_err.filename or ''}\n\n"
                "The file may be open in another application. Please close it and retry."
            )
            return
        except Exception as gen_err:
            self.error_occurred.emit("Report Error", f"Error generating reports:\n{gen_err}")
            return

        elapsed_seconds = time.time() - start_time
        stats = {
            "total_folders": len(folder_data),
            "total_files": total_discovered_files,
            "total_size": total_size_bytes,
            "elapsed": elapsed_seconds
        }
        self.inventory_finished.emit(generated_files, stats, errors)


class HashVerificationWorker(QThread):
    """
    QThread worker for batch hash computing and checksum verification.
    """
    started_verification = Signal(int)
    item_started = Signal(str, int, int, str)
    item_completed = Signal(str, dict)
    verification_cancelled = Signal(int)
    verification_completed = Signal(dict)

    def __init__(self, file_items, default_algorithm, parent=None):
        super().__init__(parent)
        self.file_items = copy.deepcopy(file_items)
        self.default_algorithm = default_algorithm
        self._cancel_event = threading.Event()
        self.ui_results = queue.SimpleQueue()
        self.ui_progress = None

    def cancel(self):
        self._cancel_event.set()

    def run(self):
        total = len(self.file_items)
        self.started_verification.emit(total)

        matched_count = 0
        mismatched_count = 0
        computed_count = 0
        error_count = 0

        for idx, item in enumerate(self.file_items, start=1):
            if self._cancel_event.is_set():
                self.verification_cancelled.emit(idx - 1)
                return

            item.update(calculated_hash="", all_hashes={}, error=None, status="pending", result="PENDING")
            fpath = Path(item["path"])
            file_id = item["id"]
            algo = item.get("algorithm") or self.default_algorithm
            expected = (item.get("expected_hash") or "").strip()

            self.item_started.emit(file_id, idx, total, item["name"])
            self.ui_progress = (file_id, idx, total, item['name'])

            try:
                if not fpath.is_file():
                    raise FileNotFoundError('File does not exist or is inaccessible')
                item['size'] = fpath.stat().st_size
                if expected and detect_hash_algorithm(expected) != algo:
                    raise ValueError("Expected hash must be hexadecimal with the correct length for " + algo)
                all_h = compute_all_file_hashes(fpath, cancel_event=self._cancel_event, algorithms=ENABLED_HASH_ALGORITHMS | {algo})
                if all_h is None:
                    self.verification_cancelled.emit(idx - 1)
                    return

                item["all_hashes"] = all_h
                calc_val = all_h.get(algo, "")
                item["calculated_hash"] = calc_val

                if expected:
                    if calc_val.lower() == expected.lower():
                        item["status"] = "match"
                        item["status_text"] = "Verified Match"
                        item["result"] = "MATCH"
                        matched_count += 1
                    else:
                        item["status"] = "mismatch"
                        item["status_text"] = "Mismatch!"
                        item["result"] = "MISMATCH"
                        mismatched_count += 1
                else:
                    item["status"] = "computed"
                    item["status_text"] = "Calculated"
                    item["result"] = "NO TARGET"
                    computed_count += 1

            except Exception as ex:
                item["status"] = "error"
                item["status_text"] = "Read Error"
                item["result"] = "ERROR"
                item["error"] = str(ex)
                error_count += 1

            self.item_completed.emit(file_id, item)
            self.ui_results.put((file_id, item))

        summary = {
            "total": total,
            "matched": matched_count,
            "mismatched": mismatched_count,
            "computed": computed_count,
            "errors": error_count
        }
        self.verification_completed.emit(summary)


class MetadataExtractionWorker(QThread):
    """
    QThread worker that processes individual files or folders (single-level or recursive),
    extracts metadata via ExifTool safely, runs pre/post forensic hash integrity verification,
    generates requested reports, and reports real-time progress without freezing the GUI.
    """
    started_extraction = Signal(int)
    file_started = Signal(str, int, int)
    file_completed = Signal(str, dict, int, int)
    file_failed = Signal(str, str, int, int)
    extraction_completed = Signal(dict)
    extraction_cancelled = Signal(int)
    extraction_error = Signal(str)

    def __init__(self, file_paths, options, parent=None):
        super().__init__(parent)
        self.file_paths = file_paths
        self.options = options
        self._cancel_event = threading.Event()

    def cancel(self):
        self._cancel_event.set()

    def run(self):
        try:
            self._run_extraction()
        except Exception as exc:
            self.extraction_error.emit(str(exc))

    def _run_extraction(self):
        total = len(self.file_paths)
        self.started_extraction.emit(total)

        successful_count = 0
        failed_count = 0
        warning_count = 0
        verified_count = 0
        batch_records = []
        generated_reports = []

        include_hashes = self.options.get("include_hashes", True)
        verify_integrity = self.options.get("verify_integrity", True)
        preserve_raw = self.options.get("preserve_raw_json", False)
        gen_html = self.options.get("gen_html", True)
        gen_txt = self.options.get("gen_txt", True)
        gen_json = self.options.get("gen_json", False)
        output_dir = self.options.get("output_dir", "")
        custom_exe = self.options.get("custom_exiftool_path", None)

        if output_dir:
            Path(output_dir).mkdir(parents=True, exist_ok=True)

        for idx, fpath in enumerate(self.file_paths, 1):
            if self._cancel_event.is_set():
                self.extraction_cancelled.emit(idx - 1)
                return

            self.file_started.emit(fpath, idx, total)

            # Step 1: Pre-extraction integrity hash (if requested)
            pre_hashes = None
            pre_hash_error = None
            if include_hashes:
                try:
                    pre_hashes = MetadataHashAdapter.calculate_integrity(fpath, cancel_event=self._cancel_event)
                except Exception as error:
                    pre_hashes = None
                    pre_hash_error = str(error)

            if self._cancel_event.is_set():
                self.extraction_cancelled.emit(idx - 1)
                return

            # Step 2: Safe ExifTool extraction
            record = ExifToolEngine.extract_metadata(fpath, custom_path=custom_exe)
            if pre_hash_error:
                record.setdefault('warnings', []).append('Pre-extraction hashing failed: ' + pre_hash_error)

            if self._cancel_event.is_set():
                self.extraction_cancelled.emit(idx - 1)
                return

            # Step 3: Post-extraction integrity hash verification (if requested)
            integrity_info = None
            if include_hashes:
                if verify_integrity and pre_hashes:
                    try:
                        integrity_info = MetadataHashAdapter.verify_integrity(pre_hashes, fpath, cancel_event=self._cancel_event)
                        if integrity_info.get("verified") and record.get("success"):
                            verified_count += 1
                    except Exception as error:
                        integrity_info = {"verified": False, "status": "HASH_FAILED", "pre": pre_hashes, "post": {}, "message": "Post-hash check error"}
                        record.setdefault('warnings', []).append('Post-extraction hashing failed: ' + str(error))
                else:
                    integrity_info = {"verified": None if pre_hashes else False, "status": "NOT_REQUESTED" if pre_hashes else "HASH_FAILED", "pre": pre_hashes or {}, "post": {}, "message": "Pre-hash only" if pre_hashes else "Hash calculation failed"}

            if self._cancel_event.is_set():
                self.extraction_cancelled.emit(idx - 1)
                return
            record["integrity_info"] = integrity_info

            if record.get("success"):
                successful_count += 1
            else:
                failed_count += 1
                self.file_failed.emit(fpath, record.get("error", "Unknown error"), idx, total)

            # Step 4: Report Exports (only when user explicitly designates auto_export)
            if output_dir and record.get("success") and self.options.get("auto_export", False):
                try:
                    sanitized_name = metadata_report_basename(record)

                    if gen_html:
                        html_path = str(Path(output_dir) / f"{sanitized_name}_Metadata.html")
                        MetadataReportExporter.export_html(html_path, record, integrity_info=integrity_info, include_hashes=include_hashes)
                        generated_reports.append(html_path)

                    if gen_txt:
                        txt_path = str(Path(output_dir) / f"{sanitized_name}_Metadata.txt")
                        MetadataReportExporter.export_txt(txt_path, record, integrity_info=integrity_info, include_hashes=include_hashes)
                        generated_reports.append(txt_path)

                    if gen_json:
                        j_rec, j_raw = MetadataReportExporter.export_json(
                            output_dir, record, integrity_info=integrity_info,
                            include_hashes=include_hashes, preserve_raw=preserve_raw
                        )
                        if j_rec: generated_reports.append(j_rec)
                        if j_raw: generated_reports.append(j_raw)
                except Exception as ex:
                    record.setdefault("warnings", []).append(f"Report export error: {ex}")

            batch_records.append(record)
            warning_count += len(record.get('warnings', []))
            self.file_completed.emit(fpath, record, idx, total)

        summary = {
            "total": total,
            "successful": successful_count,
            "failed": failed_count,
            "warnings": warning_count,
            "verified": verified_count,
            "batch_records": batch_records,
            "generated_reports": generated_reports,
            "output_dir": output_dir,
            "exiftool_version": ExifToolEngine._cached_version or "Unknown"
        }
        self.extraction_completed.emit(summary)


# ==============================================================================
# SECTION 6: MODERN SURFACE FORENSIC WORKSTATION CUSTOM UI WIDGETS
# ==============================================================================

def status_green(widget):
    """Use a readable, standard status green in either application theme."""
    current = widget
    while current is not None:
        theme = getattr(current, "current_theme", None)
        if theme in ("light", "dark"):
            return "#287a45" if theme == "light" else "#80cc93"
        current = current.parentWidget() if hasattr(current, "parentWidget") else None
    return "#80cc93"


class LucideIcons:
    """Provides crisp, scalable vector SVG icons from the Lucide icon library."""
    _SVGS = {
        "home": '<path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>',
        "shield": '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/>',
        "folder": '<path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/>',
        "search": '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
        "settings": '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
        "sun": '<circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/>',
        "moon": '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
        "compare": '<circle cx="18" cy="18" r="3"/><circle cx="6" cy="6" r="3"/><path d="M13 6h3a2 2 0 0 1 2 2v7"/><path d="M11 18H8a2 2 0 0 1-2-2V9"/>',
        "copy": '<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
        "check": '<path d="M20 6 9 17l-5-5"/>',
        "check_circle": '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>',
        "x": '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
        "x_circle": '<circle cx="12" cy="12" r="10"/><path d="m15 9-6 6"/><path d="m9 9 6 6"/>',
        "refresh": '<path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.85.83 6.72 2.24"/><path d="M21 3v6h-6"/>',
        "file_text": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>',
        "file_spreadsheet": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M8 13h2"/><path d="M14 13h2"/><path d="M8 17h2"/><path d="M14 17h2"/>',
        "file_code": '<path d="M10 12.5 8 15l2 2.5"/><path d="m14 12.5 2 2.5-2 2.5"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/>',
        "database": '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/><path d="M3 12c0 1.66 4 3 9 3s9-1.34 9-3"/>',
        "info": '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/>',
        "external_link": '<path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>',
        "play": '<polygon points="6 3 20 12 6 21 6 3"/>',
        "square": '<rect width="18" height="18" x="3" y="3" rx="2"/>',
        "trash": '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>',
        "plus": '<path d="M5 12h14"/><path d="M12 5v14"/>',
        "lock": '<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
        "cpu": '<rect width="16" height="16" x="4" y="4" rx="2"/><rect width="6" height="6" x="9" y="9" rx="1"/><path d="M15 2v2"/><path d="M15 20v2"/><path d="M2 15h2"/><path d="M2 9h2"/><path d="M20 15h2"/><path d="M20 9h2"/><path d="M9 2v2"/><path d="M9 20v2"/>',
        "terminal": '<polyline points="4 17 10 11 4 5"/><line x1="12" x2="20" y1="19" y2="19"/>',
        "activity": '<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
        "layers": '<path d="m12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/><path d="m22 12.5-8.58 3.91a2 2 0 0 1-1.66 0L2 12.5"/><path d="m22 17.5-8.58 3.91a2 2 0 0 1-1.66 0L2 17.5"/>',
        "chevron_right": '<path d="m9 18 6-6-6-6"/>',
        "eye": '<path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/>',
        "upload": '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>',
        "zap": '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
        "clock": '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
        "film": '<rect width="18" height="18" x="3" y="3" rx="2"/><path d="M7 3v18"/><path d="M3 7.5h4"/><path d="M3 12h18"/><path d="M3 16.5h4"/><path d="M17 3v18"/><path d="M17 7.5h4"/><path d="M17 16.5h4"/>',
        "music": '<path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/>',
        "file": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/>',
    }

    @classmethod
    def pixmap(cls, name, color="#3b82f6", size=20):
        inner = cls._SVGS.get(name, "")
        if not inner:
            return QPixmap()
        svg = f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{inner}</svg>'
        renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
        pm = QPixmap(size, size)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        renderer.render(p)
        p.end()
        return pm

    @classmethod
    def icon(cls, name, color="#3b82f6", size=20):
        pm = cls.pixmap(name, color=color, size=size)
        return QIcon(pm) if not pm.isNull() else QIcon()


class SidebarNavButton(QPushButton):
    """
    Modern sidebar navigation item matching Reference Image 1 & 2:
    - Pill shape with 8px radius
    - Smooth hover effect
    - Active state with bold accent background & high-contrast text
    - Lucide SVG icon on left + text
    """
    def __init__(self, text, icon_name, parent=None):
        super().__init__(parent)
        self.icon_name = icon_name
        self.nav_text = text
        self.is_active = False
        self.current_theme = "dark"
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setFixedHeight(40)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._update_appearance()

    def set_active(self, active: bool, theme: str = "dark"):
        self.is_active = active
        self.current_theme = theme
        self._update_appearance()

    def _update_appearance(self):
        if self.current_theme == "light":
            if self.is_active:
                ic_color = "#1d4ed8"
                txt_color = "#1d4ed8"
                bg_color = "#eff6ff"
                border = "1px solid #bfdbfe"
                weight = "700"
            else:
                ic_color = "#64748b"
                txt_color = "#334155"
                bg_color = "transparent"
                border = "none"
                weight = "600"
            hover_bg = "#f1f5f9"
            hover_txt = "#0f172a"
        else: # dark
            if self.is_active:
                ic_color = "#60a5fa"
                txt_color = "#ffffff"
                bg_color = "rgba(59, 130, 246, 0.22)"
                border = "1px solid rgba(59, 130, 246, 0.45)"
                weight = "700"
            else:
                ic_color = "#8e98b0"
                txt_color = "#94a3b8"
                bg_color = "transparent"
                border = "none"
                weight = "600"
            hover_bg = "#182038"
            hover_txt = "#f1f5f9"

        self.setIcon(LucideIcons.icon(self.icon_name, color=ic_color, size=18))
        self.setIconSize(QSize(18, 18))
        self.setText(f"  {self.nav_text}")
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg_color};
                border: {border};
                border-radius: 8px;
                color: {txt_color};
                text-align: left;
                padding: 8px 14px;
                font-size: 13px;
                font-weight: {weight};
            }}
            QPushButton:hover {{
                background-color: {hover_bg};
                color: {hover_txt};
            }}
        """)


def setup_label_icon(lbl: QLabel, icon_item, size=18, default_color="#AAB6C8"):
    """Display a modern Lucide SVG icon, vector PNG asset, or symbol cleanly."""
    if not icon_item:
        lbl.hide()
        return
    if lbl.parentWidget() is not None:
        lbl.show()
    else:
        # Layouts will show the label after adopting it. Showing it now creates
        # a native top-level window and flashes a title bar during startup.
        lbl.setAttribute(Qt.WA_WState_ExplicitShowHide, False)
    if isinstance(icon_item, str):
        if icon_item in LucideIcons._SVGS:
            pm = LucideIcons.pixmap(icon_item, color=default_color, size=size)
            lbl.setPixmap(pm)
            lbl.setStyleSheet("background: transparent; border: none;")
            return
        if icon_item.endswith(".png") or os.path.isfile(icon_item):
            pm = QPixmap(icon_item)
            if not pm.isNull():
                lbl.setPixmap(pm.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                lbl.setStyleSheet("background: transparent; border: none;")
                return
    elif isinstance(icon_item, QPixmap):
        lbl.setPixmap(icon_item.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        lbl.setStyleSheet("background: transparent; border: none;")
        return
    elif isinstance(icon_item, QIcon):
        lbl.setPixmap(icon_item.pixmap(size, size))
        lbl.setStyleSheet("background: transparent; border: none;")
        return

    # Text / Unicode symbol
    lbl.setText(str(icon_item))
    lbl.setStyleSheet(f"font-size: {size}px; background: transparent; border: none; color: {default_color};")


class SurfaceCard(QFrame):
    """
    Modern container card styled after Microsoft Surface Dashboard:
    Deep navy backdrop (or clean slate in light mode), subtle border, rounded 12px corners, and clean spacing.
    """
    def __init__(self, title=None, icon=None, parent=None):
        super().__init__(parent)
        self.setObjectName("SurfaceCard")
        self.setStyleSheet("""
            #SurfaceCard {
                background-color: #161b2b;
                border: 1px solid #232b43;
                border-radius: 12px;
            }
        """)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(14, 12, 14, 12)
        self.main_layout.setSpacing(8)

        self.raw_icon = icon
        self.icon_lbl = None
        self.title_lbl = None
        if title:
            header_layout = QHBoxLayout()
            header_layout.setSpacing(8)
            header_layout.setContentsMargins(0, 0, 0, 4)

            self.icon_lbl = QLabel()
            setup_label_icon(self.icon_lbl, icon, size=18)
            header_layout.addWidget(self.icon_lbl)

            self.title_lbl = QLabel(title)
            self.title_lbl.setStyleSheet("""
                font-family: 'Segoe UI', -apple-system, sans-serif;
                font-size: 13px;
                font-weight: 700;
                color: #e2e8f0;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                background: transparent;
                border: none;
            """)
            header_layout.addWidget(self.title_lbl)
            header_layout.addStretch()
            self.main_layout.addLayout(header_layout)

    def set_title(self, title):
        if hasattr(self, "title_lbl") and self.title_lbl:
            self.title_lbl.setText(title)

    def set_icon(self, icon_item, size=18):
        if self.icon_lbl:
            setup_label_icon(self.icon_lbl, icon_item, size=size)

    def apply_theme(self, theme="dark"):
        if theme == "light":
            self.setStyleSheet("""
                SurfaceCard, #SurfaceCard {
                    background-color: #ffffff;
                    border: 1px solid #CBD5E1;
                    border-radius: 12px;
                }
            """)
            if self.title_lbl:
                self.title_lbl.setStyleSheet("""
                    font-family: 'Segoe UI', -apple-system, sans-serif;
                    font-size: 13px;
                    font-weight: 700;
                    color: #0f172a;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    background: transparent;
                    border: none;
                """)
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(14)
            shadow.setColor(QColor(15, 23, 42, 25))
            shadow.setOffset(0, 2)
            self.setGraphicsEffect(shadow)
        else:
            self.setStyleSheet("""
                SurfaceCard, #SurfaceCard {
                    background-color: #161b2b;
                    border: 1px solid #232b43;
                    border-radius: 12px;
                }
            """)
            if self.title_lbl:
                self.title_lbl.setStyleSheet("""
                    font-family: 'Segoe UI', -apple-system, sans-serif;
                    font-size: 13px;
                    font-weight: 700;
                    color: #e2e8f0;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    background: transparent;
                    border: none;
                """)
            self.setGraphicsEffect(None)

        if getattr(self, "icon_lbl", None) and getattr(self, "raw_icon", None):
            ic_col = "#64748B" if theme == "light" else "#AAB6C8"
            setup_label_icon(self.icon_lbl, self.raw_icon, size=18, default_color=ic_col)


class ModernStatCard(QFrame):
    """
    Metric card matching the Hardware Monitor cards in the reference screenshot:
    Icon, title, big stat value, and subtle status/progress bar.
    """
    def __init__(self, icon_text, title, initial_val="—", subtitle="", parent=None):
        super().__init__(parent)
        self.setObjectName("ModernStatCard")
        self.setStyleSheet("""
            #ModernStatCard {
                background-color: #1c2236;
                border: 1px solid #232b43;
                border-radius: 10px;
                padding: 8px 10px;
            }
            #ModernStatCard:hover {
                border-color: #3b82f6;
                background-color: #20273d;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(2)

        top_row = QHBoxLayout()
        top_row.setSpacing(6)
        self.raw_icon = icon_text
        self.icon_lbl = QLabel()
        setup_label_icon(self.icon_lbl, icon_text, size=16)
        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0; text-transform: uppercase; background: transparent; border: none;")
        top_row.addWidget(self.icon_lbl)
        top_row.addWidget(self.title_lbl)
        top_row.addStretch()
        layout.addLayout(top_row)

        self.value_lbl = QLabel(initial_val)
        self.value_lbl.setStyleSheet("""
            font-family: 'Segoe UI', sans-serif;
            font-size: 20px;
            font-weight: 700;
            color: #f1f5f9;
            background: transparent;
            border: none;
            margin-top: 2px;
        """)
        layout.addWidget(self.value_lbl)

        self.sub_lbl = QLabel(subtitle)
        self.sub_lbl.setStyleSheet("font-size: 11px; color: #64748b; background: transparent; border: none;")
        layout.addWidget(self.sub_lbl)

        # Micro progress bar
        self.micro_bar = QProgressBar()
        self.micro_bar.setFixedHeight(4)
        self.micro_bar.setTextVisible(False)
        self.micro_bar.setStyleSheet("""
            QProgressBar {
                background-color: #111422;
                border: none;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background-color: #3b82f6;
                border-radius: 2px;
            }
        """)
        self.micro_bar.setValue(0)
        layout.addWidget(self.micro_bar)

    def set_data(self, value, subtitle=None, progress=None):
        self.value_lbl.setText(str(value))
        if subtitle is not None:
            self.sub_lbl.setText(str(subtitle))
        if progress is not None:
            self.micro_bar.setValue(int(progress))

    def apply_theme(self, theme="dark"):
        if theme == "light":
            self.setStyleSheet("""
                #ModernStatCard {
                    background-color: #ffffff;
                    border: 1px solid #CBD5E1;
                    border-radius: 10px;
                    padding: 12px;
                }
                #ModernStatCard:hover {
                    border-color: #2563eb;
                    background-color: #F8FAFC;
                }
            """)
            self.title_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #475569; text-transform: uppercase; background: transparent; border: none;")
            self.value_lbl.setStyleSheet("font-family: 'Segoe UI', sans-serif; font-size: 20px; font-weight: 700; color: #0f172a; background: transparent; border: none; margin-top: 2px;")
            self.sub_lbl.setStyleSheet("font-size: 11px; font-weight: 500; color: #475569; background: transparent; border: none;")
            self.micro_bar.setStyleSheet("""
                QProgressBar {
                    background-color: #e2e8f0;
                    border: none;
                    border-radius: 2px;
                }
                QProgressBar::chunk {
                    background-color: #2563eb;
                    border-radius: 2px;
                }
            """)
            shadow = QGraphicsDropShadowEffect(self)
            shadow.setBlurRadius(10)
            shadow.setColor(QColor(15, 23, 42, 22))
            shadow.setOffset(0, 2)
            self.setGraphicsEffect(shadow)
        else:
            self.setStyleSheet("""
                #ModernStatCard {
                    background-color: #161b2b;
                    border: 1px solid #232b43;
                    border-radius: 10px;
                    padding: 12px;
                }
                #ModernStatCard:hover {
                    border-color: #3b82f6;
                    background-color: #20273d;
                }
            """)
            self.title_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0; text-transform: uppercase; background: transparent; border: none;")
            self.value_lbl.setStyleSheet("font-family: 'Segoe UI', sans-serif; font-size: 20px; font-weight: 700; color: #f1f5f9; background: transparent; border: none; margin-top: 2px;")
            self.sub_lbl.setStyleSheet("font-size: 11px; color: #64748b; background: transparent; border: none;")
            self.micro_bar.setStyleSheet("""
                QProgressBar {
                    background-color: #111422;
                    border: none;
                    border-radius: 2px;
                }
                QProgressBar::chunk {
                    background-color: #3b82f6;
                    border-radius: 2px;
                }
            """)
            self.setGraphicsEffect(None)

        if getattr(self, "icon_lbl", None) and getattr(self, "raw_icon", None):
            if self.raw_icon == "check_circle":
                ic_col = "#287a45" if theme == "light" else "#2f7d46"
            elif self.raw_icon == "x_circle":
                ic_col = "#dc2626" if theme == "light" else "#ef4444"
            else:
                ic_col = "#64748B" if theme == "light" else "#AAB6C8"
            setup_label_icon(self.icon_lbl, self.raw_icon, size=16, default_color=ic_col)


class ProcessingSpinner(QWidget):
    """Shared three-dot pulse indicator for inventory and metadata processing."""
    def __init__(self, icon_path=None, size=24, parent=None):
        super().__init__(parent)
        self.setFixedSize(40, size)
        self.is_running = False
        self._started_at = 0.0
        self.timer = QTimer(self)
        self.timer.setInterval(33)
        self.timer.timeout.connect(self.update)
        self.setAccessibleName("Processing indicator")
        self.setAccessibleDescription("Idle")

    def start(self):
        if not self.is_running:
            self._started_at = time.monotonic()
        self.is_running = True
        self.setAccessibleDescription("Processing")
        self.timer.start()
        self.update()

    def stop(self):
        self.is_running = False
        self.timer.stop()
        self.setAccessibleDescription("Idle")
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        phase = (time.monotonic() - self._started_at) * (2 * math.pi / 1.2) if self.is_running else 0
        for index in range(3):
            pulse = (math.sin(phase - index * 0.8) + 1) / 2 if self.is_running else 0
            radius = 2.5 + 0.6 * pulse if self.is_running else 2.5
            color = QColor("#3b82f6" if self.is_running else "#8e98b0")
            color.setAlpha(int(85 + 170 * pulse) if self.is_running else 110)
            painter.setBrush(color)
            x, y = self.width() / 2 + (index - 1) * 11, self.height() / 2
            painter.drawEllipse(QRectF(x - radius, y - radius, 2 * radius, 2 * radius))



class AnimatedVerificationLoader(QWidget):
    """
    Futuristic radar/shield verification loader widget.
    Paints smooth glowing rotating rings, pulsing energy core, and status badges.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(180, 152)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.state = "idle"  # idle, verifying, match, mismatch
        self.angle = 0
        self._verification_started_at = None
        self._animation_started_at = time.monotonic()
        self.pulse = 0.0
        self.pulse_dir = 1
        self.status_text = "Integrity Engine Ready"
        self.sub_text = "Add files and click Verify Checksums"

        self.timer = QTimer(self)
        self.timer.setTimerType(Qt.TimerType.PreciseTimer)
        self.timer.timeout.connect(self._on_tick)
        self.timer.start(16)  # 60 FPS

    def _on_tick(self):
        if self.state == "verifying":
            if self._verification_started_at is None:
                self._verification_started_at = time.monotonic()
            elapsed = time.monotonic() - self._verification_started_at
            self.angle = (elapsed * 90.0) % 360.0
        else:
            self.angle = ((time.monotonic() - self._animation_started_at) * 30.0) % 360.0
        # Derive both effects from time, independently of timer jitter/file callbacks.
        self.pulse = (1.0 - math.cos((time.monotonic() - self._animation_started_at) * math.tau / 2.4)) / 2.0

        self.update()

    def set_state(self, state, status_text=None, sub_text=None):
        if state == "verifying":
            if self.state != "verifying":
                self._verification_started_at = time.monotonic()
                self.angle = 0
        else:
            self._verification_started_at = None
        self.state = state
        if status_text: self.status_text = status_text
        if sub_text: self.sub_text = sub_text
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w = self.width()
        h = self.height()
        center_x = w / 2.0
        title_height = 20 if h <= 155 else 22
        subtitle_height = 18 if h <= 155 else 20
        text_gap = 8
        text_y = h - (title_height + text_gap + subtitle_height + 12)
        graphic_bottom = text_y - 12
        center_y = (8 + graphic_bottom) / 2
        radius = min(w / 3.4, max(12, (graphic_bottom - 8) / 2))

        # Select color palette based on state
        if self.state == "verifying":
            ring_color = QColor(59, 130, 246)
            glow_color = QColor(34, 211, 238)
        elif self.state == "match":
            ring_color = QColor(16, 185, 129)
            glow_color = QColor(52, 211, 153)
        elif self.state == "mismatch":
            ring_color = QColor(239, 68, 68)
            glow_color = QColor(248, 113, 113)
        else:
            ring_color = QColor(37, 99, 235)
            glow_color = QColor(6, 182, 212)

        # 1. Ambient Background Glow
        rad_grad = QRadialGradient(center_x, center_y, radius * 1.5)
        glow_alpha = int(35 + 25 * self.pulse)
        c_glow = QColor(glow_color)
        c_glow.setAlpha(glow_alpha)
        rad_grad.setColorAt(0.0, c_glow)
        rad_grad.setColorAt(1.0, QColor(0, 0, 0, 0))
        painter.setBrush(QBrush(rad_grad))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(QPoint(int(center_x), int(center_y)), int(radius * 1.5), int(radius * 1.5))

        # 2. Outer Track Ring
        pen_track = QPen(QColor(35, 43, 67), 3)
        painter.setPen(pen_track)
        painter.setBrush(Qt.NoBrush)
        painter.drawEllipse(QPoint(int(center_x), int(center_y)), int(radius), int(radius))

        # 3. Rotating Gradient Sweep Ring
        conical = QConicalGradient(center_x, center_y, -self.angle)
        c_start = QColor(ring_color)
        c_start.setAlpha(255)
        c_end = QColor(ring_color)
        c_end.setAlpha(0)
        conical.setColorAt(0.0, c_start)
        conical.setColorAt(0.7, c_end)
        conical.setColorAt(1.0, c_end)

        sweep_pen = QPen(QBrush(conical), 4)
        sweep_pen.setCapStyle(Qt.RoundCap)
        painter.setPen(sweep_pen)
        painter.drawEllipse(QPoint(int(center_x), int(center_y)), int(radius), int(radius))

        # 4. Inner Circle Core
        inner_r = radius * 0.72
        inner_grad = QLinearGradient(center_x - inner_r, center_y - inner_r, center_x + inner_r, center_y + inner_r)
        c_inner1 = QColor(22, 27, 43)
        c_inner2 = QColor(14, 18, 29)
        inner_grad.setColorAt(0.0, c_inner1)
        inner_grad.setColorAt(1.0, c_inner2)
        painter.setBrush(QBrush(inner_grad))
        painter.setPen(QPen(QColor(35, 43, 67), 1))
        painter.drawEllipse(QPoint(int(center_x), int(center_y)), int(inner_r), int(inner_r))

        # 5. Central Modern Icon Badge
        badge_pixmap = None
        assets_dir = APP_ROOT / "assets"
        root_dir = APP_ROOT
        if self.state == "match":
            p_img = assets_dir / "check_badge.png"
            if p_img.exists(): badge_pixmap = QPixmap(str(p_img))
        elif self.state == "mismatch":
            p_img = assets_dir / "cross_badge.png"
            if p_img.exists(): badge_pixmap = QPixmap(str(p_img))
        elif self.state != "verifying":
            p_img = assets_dir / "shield.png"
            if p_img.exists(): badge_pixmap = QPixmap(str(p_img))

        if self.state == "verifying":
            globe_radius = inner_r * 0.88
            globe_rect = QRectF(
                center_x - globe_radius,
                center_y - globe_radius,
                globe_radius * 2,
                globe_radius * 2,
            )
            globe_clip = QPainterPath()
            globe_clip.addEllipse(globe_rect)
            painter.save()
            painter.setClipPath(globe_clip)
            grid_pen = QPen(QColor(72, 152, 220, 105), 1)
            painter.setPen(grid_pen)
            for latitude in (-0.62, 0.0, 0.62):
                half_width = globe_radius * math.sqrt(1 - latitude * latitude)
                half_height = max(1, globe_radius * 0.18 * math.sqrt(1 - latitude * latitude))
                latitude_y = center_y + globe_radius * latitude
                painter.drawEllipse(
                    QRectF(
                        center_x - half_width,
                        latitude_y - half_height,
                        half_width * 2,
                        half_height * 2,
                    )
                )

            for longitude in range(5):
                angle = math.radians(self.angle + longitude * 36)
                half_width = max(1, globe_radius * abs(math.cos(angle)))
                painter.drawEllipse(
                    QRectF(
                        center_x - half_width,
                        center_y - globe_radius,
                        half_width * 2,
                        globe_radius * 2,
                    )
                )

            scan_y = center_y - globe_radius + (self.angle / 360.0) * globe_radius * 2
            scan_offset = scan_y - center_y
            scan_half_width = math.sqrt(max(0, globe_radius ** 2 - scan_offset ** 2))
            beam = QLinearGradient(0, scan_y - 5, 0, scan_y + 5)
            beam.setColorAt(0.0, QColor(67, 151, 255, 0))
            beam.setColorAt(0.5, QColor(67, 151, 255, 125))
            beam.setColorAt(1.0, QColor(67, 151, 255, 0))
            painter.fillRect(
                QRectF(center_x - globe_radius, scan_y - 5, globe_radius * 2, 10),
                beam,
            )
            painter.setPen(QPen(QColor("#55aaff"), 1.5))
            painter.drawLine(
                QPointF(center_x - scan_half_width, scan_y),
                QPointF(center_x + scan_half_width, scan_y),
            )
            painter.restore()

            painter.setPen(QPen(QColor(85, 170, 255, 190), 1.5))
            painter.setBrush(Qt.NoBrush)
            painter.drawEllipse(globe_rect)
        elif badge_pixmap and not badge_pixmap.isNull():
            icon_s = int(inner_r * 1.3)
            scaled_pm = badge_pixmap.scaled(icon_s, icon_s, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            painter.drawPixmap(int(center_x - scaled_pm.width() / 2), int(center_y - scaled_pm.height() / 2), scaled_pm)
        else:
            painter.setPen(QPen(QColor(241, 245, 249)))
            font = QFont("Segoe UI", 16, QFont.Bold)
            painter.setFont(font)
            painter.drawText(
                QRect(int(center_x - inner_r), int(center_y - inner_r), int(inner_r * 2), int(inner_r * 2)),
                Qt.AlignCenter, "✓" if self.state == "match" else ("✕" if self.state == "mismatch" else "SEC")
            )

        # 6. Status Text (Crisp Typography with High Contrast)
        is_light = False
        p_widget = self.parent()
        while p_widget:
            if hasattr(p_widget, "current_theme"):
                is_light = (p_widget.current_theme == "light")
                break
            p_widget = p_widget.parent()

        status_col = QColor(15, 23, 42) if is_light else QColor(241, 245, 249)
        sub_col = QColor(71, 85, 105) if is_light else QColor(142, 152, 176)

        painter.setPen(status_col)
        font_status = QFont("Segoe UI", 10 if h <= 155 else 11, QFont.Bold)
        painter.setFont(font_status)
        painter.drawText(QRect(0, text_y, w, title_height), Qt.AlignCenter, self.status_text)

        painter.setPen(sub_col)
        font_sub = QFont("Segoe UI", 8 if h <= 155 else 9, QFont.Medium)
        painter.setFont(font_sub)
        painter.drawText(QRect(0, text_y + title_height + text_gap, w, subtitle_height), Qt.AlignCenter, self.sub_text)



class FileComparisonWorker(QThread):
    comparison_completed = Signal(str, str, str)
    comparison_failed = Signal(str)

    def __init__(self, first, second, algorithm, parent=None):
        super().__init__(parent)
        self.paths = (first, second)
        self.algorithm = algorithm
        self._cancel_event = threading.Event()

    def cancel(self):
        self._cancel_event.set()

    def run(self):
        try:
            values = []
            for path in self.paths:
                hashes = compute_all_file_hashes(path, cancel_event=self._cancel_event, algorithms={self.algorithm})
                if hashes is None:
                    return
                values.append(hashes[self.algorithm])
            if not self._cancel_event.is_set():
                self.comparison_completed.emit(self.algorithm, *values)
        except Exception as exc:
            self.comparison_failed.emit(str(exc))


class CompareTwoFilesDialog(QDialog):
    """
    Modern modal dialog for directly comparing hashes between two files.
    """
    def __init__(self, file1=None, file2=None, default_algo="SHA-256", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Compare Two Files Directly")
        self.resize(650, 480)
        assets = ensure_ui_assets()
        if assets.get("logo") and os.path.exists(assets["logo"]):
            self.setWindowIcon(application_icon())
        self.setStyleSheet("""
            QDialog {
                background-color: #0b0e17;
                color: #f1f5f9;
            }
            QLabel {
                color: #f1f5f9;
                font-family: 'Segoe UI', sans-serif;
            }
            QLineEdit {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 8px 10px;
                font-size: 12px;
            }
            QLineEdit:focus {
                border-color: #3b82f6;
            }
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 8px 14px;
                font-weight: 600;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #28304e;
                border-color: #3b82f6;
            }
            QPushButton:pressed {
                background-color: #181c2f;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        title_lbl = QLabel("Direct Two-File Hash Comparison")
        title_lbl.setStyleSheet("font-size: 16px; font-weight: 700; color: #f1f5f9;")
        layout.addWidget(title_lbl)

        desc_lbl = QLabel("Select two files to compute and compare their cryptographic checksums side-by-side.")
        desc_lbl.setStyleSheet("font-size: 12px; color: #8e98b0;")
        layout.addWidget(desc_lbl)

        # File 1 Card
        f1_card = SurfaceCard()
        f1_card.main_layout.addWidget(QLabel("FILE 1 (ORIGINAL / REFERENCE)"))
        f1_row = QHBoxLayout()
        self.f1_edit = QLineEdit()
        self.f1_edit.setPlaceholderText("Select first file...")
        self.f1_edit.setReadOnly(True)
        if file1:
            self.f1_edit.setText(str(file1))
        f1_browse = QPushButton("Browse")
        f1_browse.clicked.connect(self._browse_f1)
        f1_row.addWidget(self.f1_edit)
        f1_row.addWidget(f1_browse)
        f1_card.main_layout.addLayout(f1_row)
        self.f1_hash_lbl = QLabel("Hash: —")
        self.f1_hash_lbl.setStyleSheet("font-family: monospace; font-size: 11px; color: #60a5fa;")
        f1_card.main_layout.addWidget(self.f1_hash_lbl)
        layout.addWidget(f1_card)

        # File 2 Card
        f2_card = SurfaceCard()
        f2_card.main_layout.addWidget(QLabel("FILE 2 (TARGET / VERIFIED COPY)"))
        f2_row = QHBoxLayout()
        self.f2_edit = QLineEdit()
        self.f2_edit.setPlaceholderText("Select second file...")
        self.f2_edit.setReadOnly(True)
        if file2:
            self.f2_edit.setText(str(file2))
        f2_browse = QPushButton("Browse")
        f2_browse.clicked.connect(self._browse_f2)
        f2_row.addWidget(self.f2_edit)
        f2_row.addWidget(f2_browse)
        f2_card.main_layout.addLayout(f2_row)
        self.f2_hash_lbl = QLabel("Hash: —")
        self.f2_hash_lbl.setStyleSheet("font-family: monospace; font-size: 11px; color: #60a5fa;")
        f2_card.main_layout.addWidget(self.f2_hash_lbl)
        layout.addWidget(f2_card)

        # Controls row
        ctrl_row = QHBoxLayout()
        ctrl_row.addWidget(QLabel("Algorithm:"))
        self.algo_combo = QComboBox()
        self.algo_combo.addItems(HASH_ALGORITHMS)
        self.algo_combo.setCurrentText(default_algo)
        self.algo_combo.setStyleSheet("""
            QComboBox {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 6px 12px;
                font-size: 12px;
            }
        """)
        ctrl_row.addWidget(self.algo_combo)
        ctrl_row.addStretch()

        assets = ensure_ui_assets()
        self.compare_btn = QPushButton(" Compute & Compare")
        self.compare_btn.setIcon(QIcon(assets["compare"]))
        self.compare_btn.setIconSize(QSize(18, 18))
        self.compare_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: #ffffff;
                font-weight: 700;
                padding: 8px 18px;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover { background-color: #60a5fa; }
            QPushButton:pressed { background-color: #2563eb; }
        """)
        self.compare_btn.clicked.connect(self._do_compare)
        ctrl_row.addWidget(self.compare_btn)
        layout.addLayout(ctrl_row)

        # Result Banner
        self.result_banner = QLabel("Ready to compare")
        self.result_banner.setAlignment(Qt.AlignCenter)
        self.result_banner.setFixedHeight(45)
        self.result_banner.setStyleSheet("""
            background-color: #161b2b;
            border: 1px solid #232b43;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 700;
            color: #8e98b0;
        """)
        layout.addWidget(self.result_banner)

        # Close button
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_box.addWidget(close_btn)
        layout.addLayout(btn_box)

    def _browse_f1(self):
        f, _ = QFileDialog.getOpenFileName(self, "Select First File")
        if f:
            self.f1_edit.setText(f)

    def _browse_f2(self):
        f, _ = QFileDialog.getOpenFileName(self, "Select Second File")
        if f:
            self.f2_edit.setText(f)

    def _do_compare(self):
        paths = (self.f1_edit.text().strip(), self.f2_edit.text().strip())
        if not all(os.path.isfile(path) for path in paths):
            QMessageBox.warning(self, "Missing File", "Select two valid files to compare.")
            return
        self.compare_btn.setEnabled(False)
        self.compare_btn.setText("Comparing...")
        self.result_banner.setText("Computing checksums...")
        self.comparison_worker = FileComparisonWorker(*paths, self.algo_combo.currentText(), self)
        self.comparison_worker.comparison_completed.connect(self._on_comparison_completed)
        self.comparison_worker.comparison_failed.connect(lambda error: QMessageBox.critical(self, "Comparison Error", error))
        self.comparison_worker.finished.connect(self._comparison_finished)
        self.comparison_worker.start()

    def _on_comparison_completed(self, algorithm, first, second):
        self.f1_hash_lbl.setText(f"Hash ({algorithm}): {first}")
        self.f2_hash_lbl.setText(f"Hash ({algorithm}): {second}")
        matched = first == second
        self.result_banner.setText("MATCH CONFIRMED - Calculated hashes match" if matched else "MISMATCH DETECTED - Checksums differ")
        self.result_banner.setStyleSheet("font-weight:700; padding:12px; color:" + (("#80cc93").replace("#80cc93", status_green(self)) if matched else "#f87171"))

    def _comparison_finished(self):
        self.compare_btn.setEnabled(True)
        self.compare_btn.setText("Compute & Compare")

    def done(self, result):
        worker = getattr(self, "comparison_worker", None)
        if worker is not None and worker.isRunning():
            worker.cancel()
            QTimer.singleShot(100, lambda: self.done(result))
            return
        super().done(result)


class HashInspectorDialog(QDialog):
    """
    Modern modal dialog showing all computed hashes (MD5, SHA-1, SHA-256, SHA-512)
    for a selected file with 1-click clipboard copy buttons.
    """
    def __init__(self, file_item, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Hash Inspector — {file_item.get('name', '')}")
        self.resize(700, 480)
        assets = ensure_ui_assets()
        if assets.get("logo") and os.path.exists(assets["logo"]):
            self.setWindowIcon(application_icon())
        self.setStyleSheet("""
            QDialog {
                background-color: #0b0e17;
                color: #f1f5f9;
            }
            QLabel {
                color: #f1f5f9;
                font-family: 'Segoe UI', sans-serif;
            }
            QLineEdit {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 6px;
                color: #60a5fa;
                padding: 6px 10px;
                font-family: monospace;
                font-size: 12px;
            }
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 6px 12px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #28304e;
                border-color: #3b82f6;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(14)

        assets = ensure_ui_assets()

        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        insp_icon = QLabel()
        setup_label_icon(insp_icon, assets["inspect"], size=20)
        title = QLabel(file_item.get('name', ''))
        title.setStyleSheet("font-size: 16px; font-weight: 700;")
        title_row.addWidget(insp_icon)
        title_row.addWidget(title)
        title_row.addStretch()
        layout.addLayout(title_row)

        meta = QLabel(f"Path: {file_item.get('path', '')}\nSize: {format_size(file_item.get('size', 0))}")
        meta.setStyleSheet("font-size: 11px; color: #8e98b0;")
        layout.addWidget(meta)

        all_h = file_item.get("all_hashes") or {}

        for algo in ["MD5", "SHA-1", "SHA-256", "SHA-512"]:
            card = SurfaceCard()
            row = QHBoxLayout()
            lbl = QLabel(f"{algo}:")
            lbl.setFixedWidth(70)
            lbl.setStyleSheet("font-weight: 700; color: #e2e8f0;")
            edit = QLineEdit(all_h.get(algo, "—"))
            edit.setReadOnly(True)
            copy_btn = QPushButton(" Copy")
            copy_btn.setIcon(QIcon(assets["paste"]))
            copy_btn.setIconSize(QSize(14, 14))
            copy_btn.clicked.connect(lambda ch=False, text=edit.text(): self._copy_to_clip(text))
            row.addWidget(lbl)
            row.addWidget(edit)
            row.addWidget(copy_btn)
            card.main_layout.addLayout(row)
            layout.addWidget(card)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_box.addWidget(close_btn)
        layout.addLayout(btn_box)

    def _copy_to_clip(self, text):
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "Copied", "Checksum copied to clipboard!")


class AboutFindexDialog(QDialog):
    """Compact product identity dialog in the existing Settings area."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About FINDEX")
        self.setWindowIcon(application_icon())
        self.setMinimumWidth(440)
        self.resize(460, 390)
        light = getattr(parent, "current_theme", "dark") == "light"
        background = "#ffffff" if light else "#0f1420"
        foreground = "#172033" if light else "#edf2f9"
        secondary = "#526176" if light else "#a8b5c8"
        border = "#cbd5e1" if light else "#344158"
        surface = "#f1f5f9" if light else "#171e2e"
        self.setStyleSheet(f"""
            QDialog {{ background: {background}; }}
            QLabel {{ color: {foreground}; background: transparent;
                      font-family: 'Segoe UI'; font-size: 13px; }}
            QPushButton {{ color: {foreground}; background: {surface};
                border: 1px solid {border}; border-radius: 5px; padding: 7px 18px; }}
            QPushButton:focus, QPushButton:hover {{ border-color: #3b82f6; }}
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 20)
        layout.setSpacing(14)
        heading = QHBoxLayout()
        heading.setSpacing(12)
        logo = QLabel()
        logo.setObjectName("AboutFindexLogo")
        logo.setPixmap(QPixmap(str(APP_ROOT / "resources" / "branding" / "favicon.png")).scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        heading.addWidget(logo)
        names = QVBoxLayout()
        names.setSpacing(3)
        title = QLabel(APP_NAME)
        title.setStyleSheet(f"font-size: 22px; font-weight: 700; color: {foreground};")
        names.addWidget(title)
        names.addWidget(QLabel(APP_SUBTITLE))
        heading.addLayout(names, 1)
        layout.addLayout(heading)
        tagline = QLabel(APP_TAGLINE)
        tagline.setStyleSheet(f"color: {secondary}; font-size: 14px;")
        layout.addWidget(tagline)
        layout.addWidget(QLabel(f"Version: {APP_VERSION}"))
        description = QLabel("FINDEX is a Windows-based digital file examination utility providing file inventory generation, cryptographic hash calculation, file integrity verification and metadata examination.")
        description.setWordWrap(True)
        layout.addWidget(description)
        engine_version = getattr(parent, "exiftool_ver", None)
        engine = f"ExifTool {engine_version}" if engine_version else "ExifTool (version unavailable)"
        technical = QLabel(f"Platform: Windows\nFramework: Python / PySide6\nMetadata extraction: {engine}")
        technical.setStyleSheet(f"color: {secondary};")
        layout.addWidget(technical)
        buttons = QHBoxLayout()
        buttons.addStretch()
        close = QPushButton("Close")
        close.setDefault(True)
        close.clicked.connect(self.accept)
        buttons.addWidget(close)
        layout.addLayout(buttons)


class CompletionDialog(QDialog):
    """Inventory summary with consistent desktop buttons and outline icons."""

    def __init__(self, generated_files, stats, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inventory Complete")
        self.setMinimumWidth(480)
        self.resize(520, 370)
        self.setSizeGripEnabled(False)
        theme = getattr(parent, "current_theme", "dark")
        light = theme == "light"
        background = "#ffffff" if light else "#0f1420"
        surface = "#f8fafc" if light else "#171e2e"
        border = "#cbd5e1" if light else "#344158"
        foreground = "#172033" if light else "#edf2f9"
        muted = "#526176" if light else "#a8b5c8"
        hover = "#eaf0f7" if light else "#243047"
        pressed = "#dde6f1" if light else "#2b3952"
        success = "#287a45" if light else ("#80cc93").replace("#80cc93", status_green(self))

        # Supply several raster sizes from the existing SVG paths for HiDPI icons.
        def icon(name, color):
            result = QIcon()
            for size in (16, 20, 24, 32, 40, 48):
                result.addPixmap(LucideIcons.pixmap(name, color=color, size=size))
            return result

        self.setWindowIcon(application_icon())
        self.setStyleSheet(f"""
            QDialog {{ background: {background}; color: {foreground}; }}
            QLabel {{ color: {foreground}; background: transparent;
                      font-family: 'Segoe UI'; font-size: 13px; }}
            QFrame#completionSummary {{ background: {surface};
                border: 1px solid {border}; border-radius: 6px; }}
            QPushButton {{ background: {surface}; color: {foreground};
                border: 1px solid {border}; border-radius: 5px;
                padding: 8px 14px; font-family: 'Segoe UI';
                font-size: 13px; font-weight: 500; }}
            QPushButton:hover {{ background: {hover}; border-color: #64748b; }}
            QPushButton:pressed {{ background: {pressed}; }}
            QPushButton:focus {{ border: 1px solid #60a5fa; }}
            QPushButton#completionDone {{ background: #2563eb; color: #ffffff;
                border: 1px solid #2563eb; font-weight: 600; }}
            QPushButton#completionDone:hover {{ background: #1d4ed8;
                border-color: #1d4ed8; }}
            QPushButton#completionDone:pressed {{ background: #1e40af; }}
            QPushButton#completionDone:focus {{ border: 1px solid #93c5fd; }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 20)
        layout.setSpacing(18)

        header_row = QHBoxLayout()
        header_row.setSpacing(12)
        status_icon = QLabel()
        status_icon.setPixmap(icon("check_circle", success).pixmap(QSize(28, 28)))
        status_icon.setFixedSize(28, 28)
        header_row.addWidget(status_icon, 0, Qt.AlignTop)
        header_text = QVBoxLayout()
        header_text.setSpacing(4)
        heading = QLabel("Inventory complete")
        heading.setStyleSheet(f"color: {foreground}; font-size: 18px; font-weight: 600;")
        subtitle = QLabel("Your reports are ready to open.")
        subtitle.setStyleSheet(f"color: {muted}; font-size: 13px;")
        header_text.addWidget(heading)
        header_text.addWidget(subtitle)
        header_row.addLayout(header_text, 1)
        layout.addLayout(header_row)

        summary = QFrame()
        summary.setObjectName("completionSummary")
        grid = QGridLayout(summary)
        grid.setContentsMargins(16, 14, 16, 14)
        grid.setHorizontalSpacing(24)
        grid.setVerticalSpacing(12)
        grid.setColumnStretch(0, 1)
        for row, (title, value) in enumerate((
            ("Files processed", f"{stats.get('total_files', 0):,}"),
            ("Folder groups", f"{stats.get('total_folders', 0):,}"),
            ("Total size", format_size(stats.get("total_size", 0))),
            ("Elapsed time", f"{stats.get('elapsed', 0.0):.2f} seconds"),
        )):
            label = QLabel(title)
            label.setStyleSheet(f"color: {muted};")
            value_label = QLabel(value)
            value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            value_label.setStyleSheet(f"color: {foreground}; font-weight: 600;")
            grid.addWidget(label, row, 0)
            grid.addWidget(value_label, row, 1)
        layout.addWidget(summary)

        reports = QHBoxLayout()
        reports.setSpacing(10)
        for file_path in generated_files:
            path = str(Path(file_path).resolve())
            suffix = Path(path).suffix.lower()
            if suffix not in (".xlsx", ".html"):
                continue
            spreadsheet = suffix == ".xlsx"
            button = QPushButton("Open Excel Report" if spreadsheet else "Open HTML Report")
            button.setObjectName("openExcelReport" if spreadsheet else "openHtmlReport")
            button.setIcon(icon("file_spreadsheet" if spreadsheet else "file_code", foreground))
            button.setIconSize(QSize(18, 18))
            button.setMinimumHeight(38)
            button.setAutoDefault(False)
            button.setToolTip(path)
            button.clicked.connect(lambda checked=False, target=path: QDesktopServices.openUrl(QUrl.fromLocalFile(target)))
            reports.addWidget(button, 1)
        if reports.count():
            layout.addLayout(reports)

        divider = QFrame()
        divider.setFixedHeight(1)
        divider.setStyleSheet(f"background: {border}; border: none;")
        layout.addWidget(divider)
        footer = QHBoxLayout()
        footer.setSpacing(10)
        if generated_files:
            directory = str(Path(generated_files[0]).resolve().parent)
            folder_button = QPushButton("Open Folder")
            folder_button.setObjectName("openReportFolder")
            folder_button.setIcon(icon("folder", foreground))
            folder_button.setIconSize(QSize(18, 18))
            folder_button.setMinimumHeight(38)
            folder_button.setAutoDefault(False)
            folder_button.setToolTip(directory)
            folder_button.clicked.connect(lambda checked=False, target=directory: QDesktopServices.openUrl(QUrl.fromLocalFile(target)))
            footer.addWidget(folder_button)
        footer.addStretch()
        done = QPushButton("Done")
        done.setObjectName("completionDone")
        done.setMinimumSize(90, 38)
        done.setDefault(True)
        done.clicked.connect(self.accept)
        footer.addWidget(done)
        layout.addLayout(footer)
        done.setFocus()


class FileDropWidget(QFrame):
    """
    Drag-and-drop zone widget for files and directories with active highlight states.
    """
    path_dropped = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("FileDropWidget")
        self._theme = "dark"
        self._apply_idle_style()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(2)

        self.label_icon = QLabel()
        self.label_icon.setAlignment(Qt.AlignCenter)
        setup_label_icon(self.label_icon, "upload", size=22, default_color="#3b82f6")

        self.label_text = QLabel("Drag & drop any file or folder here to inspect")
        self.label_text.setAlignment(Qt.AlignCenter)
        self.label_text.setStyleSheet("font-size: 12px; font-weight: 600; color: #8e98b0; background: transparent; border: none;")

        self.label_sub = QLabel("Supports Images, Videos, Audio, Documents & All Files")
        self.label_sub.setAlignment(Qt.AlignCenter)
        self.label_sub.setWordWrap(True)
        self.label_sub.setStyleSheet("font-size: 10px; color: #55607c; background: transparent; border: none;")

        layout.addWidget(self.label_icon)
        layout.addWidget(self.label_text)
        layout.addWidget(self.label_sub)

    def _apply_idle_style(self):
        if self._theme == "light":
            self.setStyleSheet("""
                #FileDropWidget {
                    background-color: #f8fafc;
                    border: 2px dashed #cbd5e1;
                    border-radius: 10px;
                    padding: 8px;
                }
            """)
        else:
            self.setStyleSheet("""
                #FileDropWidget {
                    background-color: #0e121d;
                    border: 2px dashed #28334e;
                    border-radius: 10px;
                    padding: 8px;
                }
            """)

    def apply_theme(self, theme="dark"):
        self._theme = theme
        self._apply_idle_style()
        ic_col = "#2563eb" if theme == "light" else "#3b82f6"
        setup_label_icon(self.label_icon, "upload", size=26, default_color=ic_col)
        if theme == "light":
            self.label_text.setStyleSheet("font-size: 12px; font-weight: 700; color: #181E24; background: transparent; border: none;")
            self.label_sub.setStyleSheet("font-size: 11px; color: #4B5563; background: transparent; border: none;")
        else:
            self.label_text.setStyleSheet("font-size: 12px; font-weight: 600; color: #8e98b0; background: transparent; border: none;")
            self.label_sub.setStyleSheet("font-size: 10px; color: #55607c; background: transparent; border: none;")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            if self._theme == "light":
                self.setStyleSheet("""
                    #FileDropWidget {
                        background-color: #eff6ff;
                        border: 2px dashed #2563eb;
                        border-radius: 10px;
                        padding: 8px;
                    }
                """)
            else:
                self.setStyleSheet("""
                    #FileDropWidget {
                        background-color: #1a2238;
                        border: 2px dashed #3b82f6;
                        border-radius: 10px;
                        padding: 8px;
                    }
                """)

    def dragLeaveEvent(self, event):
        self._apply_idle_style()

    def dropEvent(self, event):
        self._apply_idle_style()
        urls = event.mimeData().urls()
        if urls:
            fpath = urls[0].toLocalFile()
            if fpath:
                self.path_dropped.emit(fpath)


class RawJsonDialog(QDialog):
    """
    Modal dialog displaying the original raw ExifTool JSON output exactly as received.
    """
    def __init__(self, raw_json_text, file_name="Evidence", parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Raw ExifTool JSON — {file_name}")
        self.resize(780, 560)
        assets = ensure_ui_assets()
        if assets.get("logo") and os.path.exists(assets["logo"]):
            self.setWindowIcon(application_icon())

        self.setStyleSheet("""
            QDialog {
                background-color: #0b0e17;
                color: #f1f5f9;
            }
            QTextEdit {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 8px;
                color: #93c5fd;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                padding: 10px;
            }
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #f1f5f9;
                font-weight: 600;
                padding: 8px 16px;
            }
            QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        header_lbl = QLabel(f"Original ExifTool JSON Output: {file_name}")
        header_lbl.setStyleSheet("font-size: 14px; font-weight: 700; color: #f1f5f9;")
        layout.addWidget(header_lbl)

        sub_lbl = QLabel("This view represents the exact, unformatted payload captured directly from ExifTool stdout.")
        sub_lbl.setStyleSheet("font-size: 11px; color: #8e98b0;")
        layout.addWidget(sub_lbl)

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setPlainText(raw_json_text)
        layout.addWidget(self.text_edit)

        btn_row = QHBoxLayout()
        btn_copy = QPushButton("📋 Copy to Clipboard")
        btn_copy.setCursor(QCursor(Qt.PointingHandCursor))
        btn_copy.clicked.connect(self._copy_all)

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)

        btn_row.addWidget(btn_copy)
        btn_row.addStretch()
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

    def _copy_all(self):
        QApplication.clipboard().setText(self.text_edit.toPlainText())
        QMessageBox.information(self, "Copied", "Raw JSON copied to clipboard!")


class ConfigureExifToolDialog(QDialog):
    """
    Modal configuration dialog for inspecting ExifTool binary status and selecting custom path.
    """
    configured = Signal(str)

    def __init__(self, current_path=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configure ExifTool Extraction Engine")
        self.resize(600, 320)
        assets = ensure_ui_assets()
        if assets.get("logo") and os.path.exists(assets["logo"]):
            self.setWindowIcon(application_icon())

        self.setStyleSheet("""
            QDialog {
                background-color: #0b0e17;
                color: #f1f5f9;
            }
            QLineEdit {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 8px;
            }
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(14)

        t_lbl = QLabel("ExifTool Binary Configuration")
        t_lbl.setStyleSheet("font-size: 15px; font-weight: 700;")
        layout.addWidget(t_lbl)

        d_lbl = QLabel("The Metadata Explorer requires the official Phil Harvey ExifTool executable.")
        d_lbl.setStyleSheet("font-size: 11px; color: #8e98b0;")
        layout.addWidget(d_lbl)

        path_box = SurfaceCard()
        p_row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setText(current_path or ExifToolEngine.find_executable() or "")
        btn_browse = QPushButton("Browse")
        btn_browse.clicked.connect(self._browse)
        p_row.addWidget(self.path_edit)
        p_row.addWidget(btn_browse)
        path_box.main_layout.addLayout(p_row)
        layout.addWidget(path_box)

        self.status_box = QLabel("Checking engine...")
        self.status_box.setStyleSheet("font-size: 12px; font-weight: 600; padding: 6px;")
        layout.addWidget(self.status_box)

        btn_row = QHBoxLayout()
        btn_test = QPushButton("Test Executable")
        btn_test.clicked.connect(self._test)
        btn_apply = QPushButton("Save & Apply")
        btn_apply.setStyleSheet("background-color: #3b82f6; color: #ffffff; font-weight: 700; border: none;")
        btn_apply.clicked.connect(self._apply)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_row.addWidget(btn_test)
        btn_row.addStretch()
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_apply)
        layout.addLayout(btn_row)

        self._test()

    def _browse(self):
        f, _ = QFileDialog.getOpenFileName(self, "Select exiftool.exe", "", "Executables (*.exe);;All Files (*.*)")
        if f:
            self.path_edit.setText(f)
            self._test()

    def _test(self):
        p = self.path_edit.text().strip()
        ver, ok, msg = ExifToolEngine.get_version(p if p else None)
        if ok:
            self.status_box.setText(f"✓ Operational: ExifTool version {ver} detected at:\n{msg}")
            self.status_box.setStyleSheet(("color: #80cc93; font-weight: 600;").replace("#80cc93", status_green(self)))
        else:
            self.status_box.setText(f"✕ Unavailable: {msg}")
            self.status_box.setStyleSheet("color: #f87171; font-weight: 600;")

    def _apply(self):
        p = self.path_edit.text().strip()
        self.configured.emit(p)
        self.accept()


class MediaPreviewWidget(QWidget):
    """
    Unified forensic evidence viewer for Images, Videos, Audio, and Documents.
    Provides in-application high-resolution image scaling with aspect ratio preservation,
    multimedia metadata badges, and 1-click external playback / system viewer integration.
    """
    view_metadata_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_path = None
        self.current_record = None
        self.orig_pixmap = None
        self._fit_mode = True
        self._theme = "dark"

        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # 1. Top Media Control & Info Bar
        self.top_bar = QFrame()
        self.top_bar.setStyleSheet("""
            QFrame {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 8px;
                padding: 6px 12px;
            }
        """)
        tb_layout = QHBoxLayout(self.top_bar)
        tb_layout.setContentsMargins(4, 2, 4, 2)
        tb_layout.setSpacing(10)

        self.media_title_lbl = QLabel("No evidence loaded for preview")
        self.media_title_lbl.setStyleSheet("font-size: 12px; font-weight: 700; color: #f1f5f9;")
        tb_layout.addWidget(self.media_title_lbl, 1)

        self.btn_fit = QPushButton("Fit to View")
        self.btn_fit.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_fit.setStyleSheet("""
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 5px 10px;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
        """)
        self.btn_fit.clicked.connect(self._toggle_fit_mode)
        tb_layout.addWidget(self.btn_fit)

        self.btn_open_external = QPushButton("🗖 Open in System Viewer")
        self.btn_open_external.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_open_external.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 5px 12px;
                font-size: 11px;
                font-weight: 700;
            }
            QPushButton:hover { background-color: #60a5fa; }
        """)
        self.btn_open_external.clicked.connect(self._open_external)
        tb_layout.addWidget(self.btn_open_external)

        layout.addWidget(self.top_bar)

        # 2. Stacked View Area (Index 0: Image View, Index 1: Media/Document Card, Index 2: Empty Placeholder)
        self.view_stack = QStackedWidget()

        # Page 0: Image Viewer
        self.image_scroll = QScrollArea()
        self.image_scroll.setWidgetResizable(True)
        self.image_scroll.setAlignment(Qt.AlignCenter)
        self.image_scroll.setStyleSheet("""
            QScrollArea {
                background-color: #080a10;
                border: 1px solid #1c2236;
                border-radius: 8px;
            }
        """)
        self.image_lbl = QLabel()
        self.image_lbl.setAlignment(Qt.AlignCenter)
        self.image_lbl.setStyleSheet("background: transparent; border: none;")
        self.image_scroll.setWidget(self.image_lbl)
        self.view_stack.addWidget(self.image_scroll)

        # Page 1: Media / Document Slate Card
        self.slate_card = QFrame()
        self.slate_card.setStyleSheet("""
            QFrame {
                background-color: #0e121d;
                border: 1px solid #232b43;
                border-radius: 12px;
            }
        """)
        slate_layout = QVBoxLayout(self.slate_card)
        slate_layout.setContentsMargins(24, 24, 24, 24)
        slate_layout.setSpacing(14)
        slate_layout.setAlignment(Qt.AlignCenter)

        self.slate_icon = QLabel("🎬")
        self.slate_icon.setAlignment(Qt.AlignCenter)
        self.slate_icon.setStyleSheet("font-size: 48px; background: transparent; border: none;")
        slate_layout.addWidget(self.slate_icon)

        self.slate_title = QLabel("Media Evidence")
        self.slate_title.setAlignment(Qt.AlignCenter)
        self.slate_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #f1f5f9; background: transparent; border: none;")
        slate_layout.addWidget(self.slate_title)

        self.slate_info = QLabel("Details")
        self.slate_info.setAlignment(Qt.AlignCenter)
        self.slate_info.setWordWrap(True)
        self.slate_info.setStyleSheet("font-size: 12px; color: #8e98b0; background: transparent; border: none; max-width: 500px;")
        slate_layout.addWidget(self.slate_info)

        self.slate_details_box = QLabel()
        self.slate_details_box.setAlignment(Qt.AlignCenter)
        self.slate_details_box.setStyleSheet("""
            background-color: #141a29;
            border: 1px solid #222b42;
            border-radius: 8px;
            padding: 10px 16px;
            font-size: 11px;
            color: #93c5fd;
            font-family: monospace;
        """)
        slate_layout.addWidget(self.slate_details_box)

        self.btn_slate_action = QPushButton("▶ Open in System Player")
        self.btn_slate_action.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_slate_action.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 10px 22px;
                font-size: 13px;
                font-weight: 700;
            }
            QPushButton:hover { background-color: #60a5fa; }
        """)
        self.btn_slate_action.clicked.connect(self._open_external)
        slate_layout.addWidget(self.btn_slate_action)

        self.view_stack.addWidget(self.slate_card)

        # Page 2: Empty Placeholder
        self.empty_card = QFrame()
        self.empty_card.setStyleSheet("background-color: #0b0e17; border: 1px dashed #232b43; border-radius: 8px;")
        empty_layout = QVBoxLayout(self.empty_card)
        empty_layout.setAlignment(Qt.AlignCenter)
        lbl_emp = QLabel("Select an evidence file on the left to preview it here.")
        lbl_emp.setStyleSheet("color: #64748b; font-size: 13px; font-weight: 600;")
        empty_layout.addWidget(lbl_emp)
        self.view_stack.addWidget(self.empty_card)

        layout.addWidget(self.view_stack, 1)

        self.view_stack.setCurrentIndex(2)

    def load_file(self, file_path, record=None):
        """Inspect and load file preview based on its MIME/extension category."""
        if not file_path or not os.path.exists(file_path):
            self.view_stack.setCurrentIndex(2)
            self.media_title_lbl.setText("No evidence selected")
            return

        p = Path(file_path)
        if p.is_dir():
            self.view_stack.setCurrentIndex(1)
            self.current_path = str(p)
            self.current_record = record
            self.media_title_lbl.setText(f"📁 Folder: {p.name}")
            self.slate_icon.setText("📁")
            self.slate_title.setText(f"Folder Evidence: {p.name}")
            self.slate_info.setText(f"Location: {p.resolve()}")
            self.slate_details_box.setText(f"Path: {p.resolve()}")
            self.btn_slate_action.setText("📂 Open in File Explorer")
            self.btn_fit.hide()
            return

        self.current_path = str(p)
        self.current_record = record
        ext = p.suffix.lower()
        size_str = format_size(p.stat().st_size)

        image_exts = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif", ".tiff", ".tif", ".ico", ".svg", ".dng", ".cr2", ".nef", ".arw"}
        video_exts = {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm", ".m4v", ".flv"}
        audio_exts = {".mp3", ".wav", ".aac", ".flac", ".ogg", ".m4a", ".wma"}
        doc_exts = {".pdf", ".docx", ".doc", ".xlsx", ".pptx", ".txt", ".csv", ".json", ".xml", ".html"}

        if ext in image_exts:
            pm = QPixmap(str(p))
            if not pm.isNull():
                self.orig_pixmap = pm
                self.view_stack.setCurrentIndex(0)
                self.btn_fit.show()
                self.media_title_lbl.setText(f"🖼️ {p.name}  |  {pm.width()} × {pm.height()} px  |  {size_str}")
                self.btn_open_external.setText("🗖 Open in System Viewer")
                self._render_scaled_image()
                return

        self.btn_fit.hide()
        self.orig_pixmap = None
        self.view_stack.setCurrentIndex(1)

        # Pull any available ExifTool extracted metadata
        parsed = record.get("parsed", {}) if record else {}

        if ext in video_exts:
            self.slate_icon.setText("🎬")
            self.slate_title.setText(f"Video Evidence: {p.name}")
            self.media_title_lbl.setText(f"🎬 {p.name}  |  Video  |  {size_str}")
            self.btn_open_external.setText("▶ Play Video")
            self.btn_slate_action.setText("▶ Play Video in System Player")

            duration = parsed.get("Duration", parsed.get("MediaDuration", "Available upon extraction"))
            img_size = parsed.get("ImageSize", f"{parsed.get('SourceImageWidth', '')}x{parsed.get('SourceImageHeight', '')}".strip('x') or "Available upon extraction")
            codec = parsed.get("CompressorID", parsed.get("VideoCodec", parsed.get("CompressorName", ext.upper().lstrip('.'))))
            fps = parsed.get("VideoFrameRate", parsed.get("NominalFrameRate", "—"))

            self.slate_info.setText(f"Size: {size_str}  •  Format: {ext.upper().lstrip('.')}  •  Location: {p.parent}")
            self.slate_details_box.setText(f"Duration: {duration}  |  Resolution: {img_size}  |  Codec: {codec}  |  FPS: {fps}")

        elif ext in audio_exts:
            self.slate_icon.setText("🎵")
            self.slate_title.setText(f"Audio Evidence: {p.name}")
            self.media_title_lbl.setText(f"🎵 {p.name}  |  Audio  |  {size_str}")
            self.btn_open_external.setText("▶ Play Audio")
            self.btn_slate_action.setText("▶ Play Audio in System Player")

            duration = parsed.get("Duration", "Available upon extraction")
            bitrate = parsed.get("AudioBitrate", parsed.get("Bitrate", "—"))
            channels = parsed.get("AudioChannels", parsed.get("ChannelCount", "—"))

            self.slate_info.setText(f"Size: {size_str}  •  Format: {ext.upper().lstrip('.')}  •  Location: {p.parent}")
            self.slate_details_box.setText(f"Duration: {duration}  |  Bitrate: {bitrate}  |  Channels: {channels}")

        elif ext in doc_exts:
            self.slate_icon.setText("📄")
            self.slate_title.setText(f"Document Evidence: {p.name}")
            self.media_title_lbl.setText(f"📄 {p.name}  |  Document  |  {size_str}")
            self.btn_open_external.setText("📄 Open Document")
            self.btn_slate_action.setText("📄 Open Document in System Reader")

            pages = parsed.get("PageCount", parsed.get("Pages", "—"))
            creator = parsed.get("Creator", parsed.get("Author", parsed.get("Producer", "—")))

            self.slate_info.setText(f"Size: {size_str}  •  Format: {ext.upper().lstrip('.')}  •  Location: {p.parent}")
            self.slate_details_box.setText(f"Pages: {pages}  |  Creator/Author: {creator}  |  File Type: {parsed.get('FileType', ext.upper())}")

        else:
            self.slate_icon.setText("📁")
            self.slate_title.setText(f"File Evidence: {p.name}")
            self.media_title_lbl.setText(f"📁 {p.name}  |  {size_str}")
            self.btn_open_external.setText("Open File")
            self.btn_slate_action.setText("Open File in Default Application")
            self.slate_info.setText(f"Size: {size_str}  •  Location: {p.parent}")
            self.slate_details_box.setText(f"Type: {parsed.get('MIMEType', 'Binary / Data')}")

    def _render_scaled_image(self):
        if not self.orig_pixmap:
            return
        if self._fit_mode:
            vp = self.image_scroll.viewport().size()
            target_w = max(100, vp.width() - 20)
            target_h = max(100, vp.height() - 20)
            scaled = self.orig_pixmap.scaled(target_w, target_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_lbl.setPixmap(scaled)
        else:
            self.image_lbl.setPixmap(self.orig_pixmap)

    def _toggle_fit_mode(self):
        self._fit_mode = not self._fit_mode
        self.btn_fit.setText("100% Actual" if self._fit_mode else "Fit to View")
        self._render_scaled_image()

    def _open_external(self):
        if self.current_path and os.path.exists(self.current_path):
            open_evidence_file(self, self.current_path)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.view_stack.currentIndex() == 0:
            self._render_scaled_image()

    def apply_theme(self, theme="dark"):
        self._theme = theme
        if theme == "light":
            self.top_bar.setStyleSheet("QFrame { background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 6px 12px; }")
            self.media_title_lbl.setStyleSheet("font-size: 12px; font-weight: 700; color: #0f172a;")
            self.slate_card.setStyleSheet("QFrame { background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 12px; }")
            self.slate_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #0f172a; background: transparent; border: none;")
            self.slate_info.setStyleSheet("font-size: 12px; color: #64748b; background: transparent; border: none;")
            self.slate_details_box.setStyleSheet("background-color: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 16px; font-size: 11px; color: #1e40af; font-family: monospace;")
            self.image_scroll.setStyleSheet("QScrollArea { background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; }")
            self.btn_fit.setStyleSheet("""
                QPushButton {
                    background-color: #f1f5f9;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    color: #0f172a;
                    padding: 5px 10px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover { background-color: #e2e8f0; border-color: #2563eb; }
            """)
        else:
            self.top_bar.setStyleSheet("QFrame { background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px; padding: 6px 12px; }")
            self.media_title_lbl.setStyleSheet("font-size: 12px; font-weight: 700; color: #f1f5f9;")
            self.slate_card.setStyleSheet("QFrame { background-color: #0e121d; border: 1px solid #232b43; border-radius: 12px; }")
            self.slate_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #f1f5f9; background: transparent; border: none;")
            self.slate_info.setStyleSheet("font-size: 12px; color: #8e98b0; background: transparent; border: none;")
            self.slate_details_box.setStyleSheet("background-color: #141a29; border: 1px solid #222b42; border-radius: 8px; padding: 10px 16px; font-size: 11px; color: #93c5fd; font-family: monospace;")
            self.image_scroll.setStyleSheet("QScrollArea { background-color: #080a10; border: 1px solid #1c2236; border-radius: 8px; }")
            self.btn_fit.setStyleSheet("""
                QPushButton {
                    background-color: #1e243b;
                    border: 1px solid #2e3654;
                    border-radius: 6px;
                    color: #f1f5f9;
                    padding: 5px 10px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
            """)


class StatCardShim:
    """Lightweight backward-compatible shim for legacy KPI stat card calls."""
    def __init__(self, pill_widget=None, default_val="0"):
        self.pill_widget = pill_widget
        self.value_lbl = QLabel(str(default_val))

    def set_data(self, val, subtitle=None, progress=None):
        self.value_lbl.setText(str(val))
        if self.pill_widget:
            name = self.pill_widget.objectName()
            if "Integrity" in name:
                if "Verified" in str(val):
                    self.pill_widget.setText("✓ Integrity Verified")
                    self.pill_widget.setStyleSheet(("color: #80cc93; background: rgba(128, 204, 147, 0.15); border: 1px solid rgba(128, 204, 147, 0.3); border-radius: 6px; padding: 3px 8px; font-weight: 700; font-size: 11px;").replace("#80cc93", status_green(self)))
                elif str(val) in ("—", "N/A"):
                    self.pill_widget.setText("Integrity: N/A")
                    self.pill_widget.setStyleSheet("color: #8e98b0; background: rgba(142, 152, 176, 0.12); border: 1px solid rgba(142, 152, 176, 0.25); border-radius: 6px; padding: 3px 8px; font-weight: 600; font-size: 11px;")
                else:
                    self.pill_widget.setText(f"Integrity: {val}")
            elif "Count" in name:
                self.pill_widget.setText(f"{val} Files")
                self.pill_widget.setStyleSheet("color: #93c5fd; background: rgba(59, 130, 246, 0.12); border: 1px solid rgba(59, 130, 246, 0.25); border-radius: 6px; padding: 3px 8px; font-weight: 700; font-size: 11px;")



DOC_EXTENSIONS = DOCUMENT_EXTENSIONS | {".json", ".xml", ".html"}


class DynamicPreviewGalleryViewer(SurfaceCard):
    """
    Dynamic Preview / Gallery Viewer for Column 2 of Metadata Explorer.
    Seamlessly adapts between Single File mode (large preview, aspect-ratio preservation,
    zoom/fit controls, metadata specifications) and Folder mode (responsive thumbnail gallery,
    folder statistics, filtering, pagination, and enlarged in-gallery inspection).
    """
    file_selected = Signal(str)

    def __init__(self, parent=None):
        super().__init__(title="Evidence Preview", icon="eye", parent=parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumHeight(240)

        self.current_mode = "empty"  # "empty", "single", "folder"
        self.current_theme = "dark"
        self.current_file_path = None
        self.current_folder_path = None
        self.current_record = None

        # Single file image state
        self._orig_pixmap = None
        self._zoom_factor = 1.0
        self._fit_mode = True

        # Folder gallery state
        self.folder_files = []
        self.folder_images = []
        self.folder_videos = []
        self.folder_docs = []
        self.folder_others = []
        self.filtered_files = []
        self.current_page = 0
        self.items_per_page = 12
        self.selected_item_path = None
        self.enlarged_image_index = -1
        self._thumb_cache = {}

        self._build_viewer_ui()
        self.apply_theme("dark")

    def _build_viewer_ui(self):
        self.stack = QStackedWidget()

        # ================= PAGE 0: EMPTY STATE =================
        self.page_empty = QWidget()
        emp_layout = QVBoxLayout(self.page_empty)
        emp_layout.setAlignment(Qt.AlignCenter)
        emp_layout.setSpacing(10)

        self.lbl_empty_icon = QLabel()
        self.lbl_empty_icon.setAlignment(Qt.AlignCenter)
        setup_label_icon(self.lbl_empty_icon, "eye", size=36, default_color="#64748B")
        emp_layout.addWidget(self.lbl_empty_icon)

        self.lbl_empty_title = QLabel("No Evidence Loaded")
        self.lbl_empty_title.setAlignment(Qt.AlignCenter)
        self.lbl_empty_title.setStyleSheet("font-size: 13px; font-weight: 700; color: #8e98b0;")
        emp_layout.addWidget(self.lbl_empty_title)

        self.lbl_empty_sub = QLabel("Select a single file or directory above\nor drag and drop evidence into the target zone.")
        self.lbl_empty_sub.setAlignment(Qt.AlignCenter)
        self.lbl_empty_sub.setStyleSheet("font-size: 11px; color: #55607c;")
        emp_layout.addWidget(self.lbl_empty_sub)
        self.stack.addWidget(self.page_empty)

        # ================= PAGE 1: SINGLE FILE PREVIEW =================
        self.page_single = QWidget()
        single_layout = QVBoxLayout(self.page_single)
        single_layout.setContentsMargins(0, 0, 0, 0)
        single_layout.setSpacing(6)

        # 1. Single Top Control Bar
        self.single_top_bar = QFrame()
        self.single_top_bar.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        stb_layout = QHBoxLayout(self.single_top_bar)
        stb_layout.setContentsMargins(8, 4, 8, 4)
        stb_layout.setSpacing(6)

        self.single_title_lbl = QLabel("file_name.ext")
        self.single_title_lbl.setStyleSheet("font-size: 12px; font-weight: 700; color: #f1f5f9;")
        stb_layout.addWidget(self.single_title_lbl, 1)

        self.btn_fit = QPushButton("Fit to View")
        self.btn_fit.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_fit.clicked.connect(self._toggle_single_fit_mode)
        stb_layout.addWidget(self.btn_fit)

        self.btn_zoom_out = QPushButton("-")
        self.btn_zoom_out.setFixedSize(26, 26)
        self.btn_zoom_out.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_zoom_out.clicked.connect(lambda: self._adjust_zoom(-0.2))
        stb_layout.addWidget(self.btn_zoom_out)

        self.btn_zoom_in = QPushButton("+")
        self.btn_zoom_in.setFixedSize(26, 26)
        self.btn_zoom_in.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_zoom_in.clicked.connect(lambda: self._adjust_zoom(0.2))
        stb_layout.addWidget(self.btn_zoom_in)

        self.btn_open_external = QPushButton("Open")
        self.btn_open_external.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_open_external.clicked.connect(self._open_current_external)
        stb_layout.addWidget(self.btn_open_external)

        single_layout.addWidget(self.single_top_bar)

        # 2. Viewport Stack (Image Scroll Area vs Non-Image Slate)
        self.single_view_stack = QStackedWidget()

        # Image view scroll
        self.single_scroll = QScrollArea()
        self.single_scroll.setWidgetResizable(True)
        self.single_scroll.setAlignment(Qt.AlignCenter)
        self.single_scroll.setStyleSheet("QScrollArea { background-color: #080a10; border: 1px solid #1c2236; border-radius: 8px; }")

        self.single_img_lbl = QLabel()
        self.single_img_lbl.setAlignment(Qt.AlignCenter)
        self.single_img_lbl.setStyleSheet("background: transparent; border: none;")
        self.single_scroll.setWidget(self.single_img_lbl)
        self.single_view_stack.addWidget(self.single_scroll)

        # Non-image slate card
        self.non_img_card = QFrame()
        self.non_img_card.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        nic_layout = QVBoxLayout(self.non_img_card)
        nic_layout.setAlignment(Qt.AlignCenter)
        nic_layout.setSpacing(10)

        self.non_img_icon = QLabel()
        self.non_img_icon.setAlignment(Qt.AlignCenter)
        setup_label_icon(self.non_img_icon, "file", size=48, default_color="#64748B")
        nic_layout.addWidget(self.non_img_icon)

        self.non_img_title = QLabel("Document / Media")
        self.non_img_title.setAlignment(Qt.AlignCenter)
        self.non_img_title.setStyleSheet("font-size: 14px; font-weight: 700; color: #f1f5f9;")
        nic_layout.addWidget(self.non_img_title)

        self.non_img_desc = QLabel("Standard File")
        self.non_img_desc.setAlignment(Qt.AlignCenter)
        self.non_img_desc.setStyleSheet("font-size: 11px; color: #8e98b0;")
        nic_layout.addWidget(self.non_img_desc)

        self.single_view_stack.addWidget(self.non_img_card)
        single_layout.addWidget(self.single_view_stack, 1)

        # 3. Bottom Specs Bar
        self.single_specs_frame = QFrame()
        self.single_specs_frame.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        ssf_layout = QHBoxLayout(self.single_specs_frame)
        ssf_layout.setContentsMargins(10, 6, 10, 6)
        ssf_layout.setSpacing(10)

        self.single_badge_lbl = QLabel("FORMAT")
        self.single_badge_lbl.setStyleSheet("font-size: 10px; font-weight: 800; color: #38bdf8; background: rgba(56, 189, 248, 0.15); border-radius: 4px; padding: 2px 6px;")
        ssf_layout.addWidget(self.single_badge_lbl)

        self.single_specs_lbl = QLabel("1920 x 1080 px  •  2.4 MB  •  2025-01-27 17:24:48")
        self.single_specs_lbl.setStyleSheet("font-size: 11px; color: #94a3b8; font-weight: 500;")
        ssf_layout.addWidget(self.single_specs_lbl, 1)

        single_layout.addWidget(self.single_specs_frame)
        self.stack.addWidget(self.page_single)

        # ================= PAGE 2: FOLDER GALLERY GRID =================
        self.page_gallery = QWidget()
        gal_layout = QVBoxLayout(self.page_gallery)
        gal_layout.setContentsMargins(0, 0, 0, 0)
        gal_layout.setSpacing(6)

        # 1. Gallery Header Info & Filter Bar
        self.gal_header_frame = QFrame()
        self.gal_header_frame.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        ghf_layout = QVBoxLayout(self.gal_header_frame)
        ghf_layout.setContentsMargins(10, 6, 10, 6)
        ghf_layout.setSpacing(4)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        self.gal_folder_name_lbl = QLabel("Folder: /evidence")
        self.gal_folder_name_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #f1f5f9;")
        top_row.addWidget(self.gal_folder_name_lbl, 1)

        self.gal_filter_combo = QComboBox()
        self.gal_filter_combo.addItems(["Images Only", "All Files"])
        self.gal_filter_combo.currentIndexChanged.connect(self._on_gallery_filter_changed)
        top_row.addWidget(self.gal_filter_combo)
        ghf_layout.addLayout(top_row)

        self.gal_stats_lbl = QLabel("Total: 0  |  Images: 0  |  Videos: 0  |  Other: 0")
        self.gal_stats_lbl.setStyleSheet("font-size: 10px; color: #8e98b0; font-weight: 500;")
        ghf_layout.addWidget(self.gal_stats_lbl)
        gal_layout.addWidget(self.gal_header_frame)

        # 2. Thumbnail Grid Scroll Area
        self.gal_scroll = QScrollArea()
        self.gal_scroll.setWidgetResizable(True)
        self.gal_scroll.setStyleSheet("QScrollArea { background-color: #080a10; border: 1px solid #1c2236; border-radius: 8px; }")

        self.gal_grid_container = QWidget()
        self.gal_grid_layout = QGridLayout(self.gal_grid_container)
        self.gal_grid_layout.setContentsMargins(8, 8, 8, 8)
        self.gal_grid_layout.setSpacing(8)
        self.gal_scroll.setWidget(self.gal_grid_container)
        gal_layout.addWidget(self.gal_scroll, 1)

        # 3. Gallery Pagination Bar
        self.gal_nav_frame = QFrame()
        self.gal_nav_frame.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        gnf_layout = QHBoxLayout(self.gal_nav_frame)
        gnf_layout.setContentsMargins(8, 4, 8, 4)
        gnf_layout.setSpacing(8)

        self.btn_gal_prev = QPushButton("Previous")
        self.btn_gal_prev.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_gal_prev.clicked.connect(lambda: self._change_gallery_page(-1))
        gnf_layout.addWidget(self.btn_gal_prev)

        self.gal_page_lbl = QLabel("Page 1 of 1 (0 items)")
        self.gal_page_lbl.setAlignment(Qt.AlignCenter)
        self.gal_page_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0;")
        gnf_layout.addWidget(self.gal_page_lbl, 1)

        self.btn_gal_next = QPushButton("Next")
        self.btn_gal_next.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_gal_next.clicked.connect(lambda: self._change_gallery_page(1))
        gnf_layout.addWidget(self.btn_gal_next)

        gal_layout.addWidget(self.gal_nav_frame)
        self.stack.addWidget(self.page_gallery)

        # ================= PAGE 3: ENLARGED IMAGE VIEW (WITHIN GALLERY) =================
        self.page_enlarged = QWidget()
        enl_layout = QVBoxLayout(self.page_enlarged)
        enl_layout.setContentsMargins(0, 0, 0, 0)
        enl_layout.setSpacing(6)

        # 1. Enlarged View Navigation Toolbar
        self.enl_top_bar = QFrame()
        self.enl_top_bar.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        etb_layout = QHBoxLayout(self.enl_top_bar)
        etb_layout.setContentsMargins(8, 4, 8, 4)
        etb_layout.setSpacing(6)

        self.btn_back_to_grid = QPushButton("Back to Grid")
        self.btn_back_to_grid.setIcon(LucideIcons.icon("arrow_left", color="#3b82f6", size=14))
        self.btn_back_to_grid.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_back_to_grid.clicked.connect(self._return_to_gallery_grid)
        etb_layout.addWidget(self.btn_back_to_grid)

        self.enl_counter_lbl = QLabel("Image 1 of 1")
        self.enl_counter_lbl.setAlignment(Qt.AlignCenter)
        self.enl_counter_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #f1f5f9;")
        etb_layout.addWidget(self.enl_counter_lbl, 1)

        self.btn_enl_prev = QPushButton("Previous")
        self.btn_enl_prev.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_enl_prev.clicked.connect(lambda: self._navigate_enlarged_image(-1))
        etb_layout.addWidget(self.btn_enl_prev)

        self.btn_enl_next = QPushButton("Next")
        self.btn_enl_next.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_enl_next.clicked.connect(lambda: self._navigate_enlarged_image(1))
        etb_layout.addWidget(self.btn_enl_next)

        enl_layout.addWidget(self.enl_top_bar)

        # 2. Enlarged Image Scroll Area
        self.enl_scroll = QScrollArea()
        self.enl_scroll.setWidgetResizable(True)
        self.enl_scroll.setAlignment(Qt.AlignCenter)
        self.enl_scroll.setStyleSheet("QScrollArea { background-color: #080a10; border: 1px solid #1c2236; border-radius: 8px; }")

        self.enl_img_lbl = QLabel()
        self.enl_img_lbl.setAlignment(Qt.AlignCenter)
        self.enl_img_lbl.setStyleSheet("background: transparent; border: none;")
        self.enl_scroll.setWidget(self.enl_img_lbl)
        enl_layout.addWidget(self.enl_scroll, 1)

        # 3. Enlarged Image Specs Bar
        self.enl_specs_frame = QFrame()
        self.enl_specs_frame.setStyleSheet("background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px;")
        esf_layout = QHBoxLayout(self.enl_specs_frame)
        esf_layout.setContentsMargins(10, 6, 10, 6)
        esf_layout.setSpacing(10)

        self.enl_badge_lbl = QLabel("IMAGE")
        self.enl_badge_lbl.setStyleSheet("font-size: 10px; font-weight: 800; color: #38bdf8; background: rgba(56, 189, 248, 0.15); border-radius: 4px; padding: 2px 6px;")
        esf_layout.addWidget(self.enl_badge_lbl)

        self.enl_specs_lbl = QLabel("photo.jpg  •  1920 x 1080 px  •  1.5 MB")
        self.enl_specs_lbl.setStyleSheet("font-size: 11px; color: #94a3b8; font-weight: 500;")
        esf_layout.addWidget(self.enl_specs_lbl, 1)

        enl_layout.addWidget(self.enl_specs_frame)
        self.stack.addWidget(self.page_enlarged)

        self.main_layout.addWidget(self.stack)
        self.stack.setCurrentIndex(0)

    # ================= LOADING & MODE DISPATCH =================
    def load_file(self, file_path, record=None):
        """Display single file preview or dispatch to folder mode if directory."""
        if not file_path or not os.path.exists(file_path):
            self.current_mode = "empty"
            self.current_file_path = None
            self.current_record = None
            self.stack.setCurrentIndex(0)
            self.set_title("Evidence Preview")
            return

        p = Path(file_path)
        if p.is_dir():
            self.load_folder(str(p))
            return

        self.current_mode = "single"
        self.current_file_path = str(p)
        self.current_record = record
        self.set_title("Single File Preview")

        ext = p.suffix.lower()
        stat = p.stat()
        size_str = format_size(stat.st_size)
        mtime_str = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")

        # Check if file is image
        if ext in IMAGE_EXTENSIONS:
            self._fit_mode = True
            self._zoom_factor = 1.0
            pm = QPixmap(str(p))
            if not pm.isNull():
                self._orig_pixmap = pm
                self.single_view_stack.setCurrentIndex(0)
                self.btn_fit.show()
                self.btn_zoom_in.show()
                self.btn_zoom_out.show()

                dims = f"{pm.width()} × {pm.height()} px"
                self.single_title_lbl.setText(p.name)
                self.single_title_lbl.setToolTip(str(p.resolve()))
                self.single_badge_lbl.setText(f"{ext.upper().lstrip('.')} IMAGE")
                self.single_specs_lbl.setText(f"{dims}  •  {size_str}  •  {mtime_str}")
                self._update_single_image_display()
            else:
                self._show_non_image_preview(p, ext, size_str, mtime_str)
        else:
            self._show_non_image_preview(p, ext, size_str, mtime_str)

        self.stack.setCurrentIndex(1)

    def _show_non_image_preview(self, p, ext, size_str, mtime_str):
        self._orig_pixmap = None
        self.single_view_stack.setCurrentIndex(1)
        self.btn_fit.hide()
        self.btn_zoom_in.hide()
        self.btn_zoom_out.hide()

        self.single_title_lbl.setText(p.name)
        self.single_title_lbl.setToolTip(str(p.resolve()))

        if ext in VIDEO_EXTENSIONS:
            badge_text = f"{ext.upper().lstrip('.')} VIDEO"
            icon_name = "film"
            desc = "Video Evidence File"
        elif ext in AUDIO_EXTENSIONS:
            badge_text = f"{ext.upper().lstrip('.')} AUDIO"
            icon_name = "music"
            desc = "Audio Evidence File"
        elif ext in DOC_EXTENSIONS:
            badge_text = f"{ext.upper().lstrip('.')} DOCUMENT"
            icon_name = "file_text"
            desc = "Structured Document / Data"
        elif ext in ARCHIVE_EXTENSIONS:
            badge_text = f"{ext.upper().lstrip('.')} ARCHIVE"
            icon_name = "package"
            desc = "Compressed Archive Container"
        else:
            badge_text = f"{ext.upper().lstrip('.') or 'BIN'} FILE"
            icon_name = "file"
            desc = "Binary / Evidence Data File"

        self.single_badge_lbl.setText(badge_text)
        self.single_specs_lbl.setText(f"{size_str}  •  {mtime_str}")
        self.non_img_title.setText(p.name)
        self.non_img_desc.setText(f"{desc} ({size_str})")
        setup_label_icon(self.non_img_icon, icon_name, size=48, default_color="#64748B" if self.current_theme == "light" else "#AAB6C8")

    def _update_single_image_display(self):
        if not self._orig_pixmap or self._orig_pixmap.isNull():
            return
        if self._fit_mode:
            viewport_size = self.single_scroll.viewport().size()
            vw = max(100, viewport_size.width() - 10)
            vh = max(100, viewport_size.height() - 10)
            scaled = self._orig_pixmap.scaled(vw, vh, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.single_img_lbl.setPixmap(scaled)
        else:
            w = int(self._orig_pixmap.width() * self._zoom_factor)
            h = int(self._orig_pixmap.height() * self._zoom_factor)
            scaled = self._orig_pixmap.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.single_img_lbl.setPixmap(scaled)

    def _toggle_single_fit_mode(self):
        self._fit_mode = not self._fit_mode
        self.btn_fit.setText("Zoom 100%" if self._fit_mode else "Fit to View")
        if not self._fit_mode:
            self._zoom_factor = 1.0
        self._update_single_image_display()

    def _adjust_zoom(self, delta):
        self._fit_mode = False
        self.btn_fit.setText("Fit to View")
        self._zoom_factor = max(0.2, min(5.0, self._zoom_factor + delta))
        self._update_single_image_display()

    def _open_current_external(self):
        target = self.current_file_path or self.current_folder_path
        if target and os.path.exists(target):
            open_evidence_file(self, target)

    # ================= FOLDER GALLERY MODE =================
    def load_folder(self, folder_path, recursive=False):
        """Scans the specified folder and displays the responsive thumbnail gallery."""
        if not folder_path or not os.path.isdir(folder_path):
            self.load_file(None)
            return

        self.current_mode = "folder"
        self.current_folder_path = folder_path
        self.set_title("Folder Gallery")

        root = Path(folder_path)
        all_files = []
        try:
            if recursive:
                for current, _, files in evidence_walk(root):
                    for name in files:
                        p = Path(current) / name
                        if p.is_file() and not p.name.startswith("."):
                            all_files.append(p)
            else:
                for p in root.iterdir():
                    if p.is_file() and not p.name.startswith("."):
                        all_files.append(p)
        except Exception as e:
            QMessageBox.warning(self, "Folder Preview", f"The folder preview is incomplete: {e}")

        self.folder_files = all_files
        self.folder_images = [f for f in all_files if f.suffix.lower() in IMAGE_EXTENSIONS]
        self.folder_videos = [f for f in all_files if f.suffix.lower() in VIDEO_EXTENSIONS]
        self.folder_docs = [f for f in all_files if f.suffix.lower() in DOC_EXTENSIONS]
        self.folder_others = [f for f in all_files if f not in self.folder_images and f not in self.folder_videos and f not in self.folder_docs]

        # Update stats
        folder_disp = root.name or str(root)
        self.gal_folder_name_lbl.setText(f"Folder: {folder_disp}")
        self.gal_folder_name_lbl.setToolTip(str(root.resolve()))
        self.gal_stats_lbl.setText(
            f"Total: {len(all_files)} files  |  Images: {len(self.folder_images)}  |  Videos: {len(self.folder_videos)}  |  Docs/Other: {len(self.folder_docs) + len(self.folder_others)}"
        )

        self._filter_and_render_gallery()
        self.stack.setCurrentIndex(2)

    def _on_gallery_filter_changed(self, idx):
        self.current_page = 0
        self._filter_and_render_gallery()

    def _filter_and_render_gallery(self):
        filter_type = self.gal_filter_combo.currentText()
        if filter_type == "Images Only":
            self.filtered_files = self.folder_images
        else:
            self.filtered_files = self.folder_images + self.folder_videos + self.folder_docs + self.folder_others

        self._render_gallery_grid()

    def _render_gallery_grid(self):
        # Clear existing layout items
        while self.gal_grid_layout.count():
            item = self.gal_grid_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        total_items = len(self.filtered_files)
        total_pages = max(1, math.ceil(total_items / self.items_per_page))
        self.current_page = max(0, min(self.current_page, total_pages - 1))

        start = self.current_page * self.items_per_page
        end = min(start + self.items_per_page, total_items)
        page_items = self.filtered_files[start:end]

        # Update pagination controls
        self.btn_gal_prev.setEnabled(self.current_page > 0)
        self.btn_gal_next.setEnabled(self.current_page < total_pages - 1)
        self.gal_page_lbl.setText(f"Page {self.current_page + 1} of {total_pages} (Showing {start + 1 if total_items > 0 else 0}–{end} of {total_items})")

        # Responsive columns (3 columns in standard width)
        columns = 3
        for i, fpath in enumerate(page_items):
            card = self._create_thumbnail_card(fpath)
            row = i // columns
            col = i % columns
            self.gal_grid_layout.addWidget(card, row, col)

    def _create_thumbnail_card(self, fpath: Path):
        card = QFrame()
        card.setFixedSize(100, 115)
        card.setCursor(QCursor(Qt.PointingHandCursor))

        is_selected = (str(fpath) == str(self.selected_item_path))
        theme_bg = "#ffffff" if self.current_theme == "light" else "#141a29"
        border_col = "#2563eb" if is_selected else ("#cbd5e1" if self.current_theme == "light" else "#222b42")

        card.setStyleSheet(f"""
            QFrame {{
                background-color: {theme_bg};
                border: {'2px solid #3b82f6' if is_selected else f'1px solid {border_col}'};
                border-radius: 6px;
            }}
            QFrame:hover {{
                border-color: #3b82f6;
            }}
        """)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)
        layout.setAlignment(Qt.AlignCenter)

        thumb_lbl = QLabel()
        thumb_lbl.setFixedSize(92, 75)
        thumb_lbl.setAlignment(Qt.AlignCenter)
        thumb_lbl.setStyleSheet("background: transparent; border: none;")

        ext = fpath.suffix.lower()
        if ext in IMAGE_EXTENSIONS:
            # Check thumbnail cache
            p_str = str(fpath)
            if p_str in self._thumb_cache:
                thumb_lbl.setPixmap(self._thumb_cache[p_str])
            else:
                pm = QPixmap(p_str)
                if not pm.isNull():
                    scaled = pm.scaled(90, 72, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self._thumb_cache[p_str] = scaled
                    thumb_lbl.setPixmap(scaled)
                else:
                    setup_label_icon(thumb_lbl, "image", size=32, default_color="#64748B")
        elif ext in VIDEO_EXTENSIONS:
            setup_label_icon(thumb_lbl, "film", size=32, default_color="#64748B")
        elif ext in AUDIO_EXTENSIONS:
            setup_label_icon(thumb_lbl, "music", size=32, default_color="#64748B")
        elif ext in DOC_EXTENSIONS:
            setup_label_icon(thumb_lbl, "file_text", size=32, default_color="#64748B")
        else:
            setup_label_icon(thumb_lbl, "file", size=32, default_color="#64748B")

        layout.addWidget(thumb_lbl)

        # Name label
        name_lbl = QLabel(fpath.name)
        name_lbl.setAlignment(Qt.AlignCenter)
        name_color = "#0f172a" if self.current_theme == "light" else "#f1f5f9"
        name_lbl.setStyleSheet(f"font-size: 10px; font-weight: 600; color: {name_color}; background: transparent; border: none;")
        name_lbl.setToolTip(f"{fpath.name} ({format_size(fpath.stat().st_size)})")

        fm = QFontMetrics(name_lbl.font())
        elided = fm.elidedText(fpath.name, Qt.ElideMiddle, 88)
        name_lbl.setText(elided)
        layout.addWidget(name_lbl)

        # Click handler
        card.mousePressEvent = lambda e, p=fpath: self._on_thumbnail_clicked(p)
        return card

    def _on_thumbnail_clicked(self, fpath: Path):
        self.selected_item_path = str(fpath)
        self.file_selected.emit(str(fpath))

        ext = fpath.suffix.lower()
        if ext in IMAGE_EXTENSIONS and fpath in self.folder_images:
            # Open in enlarged in-gallery view!
            self.enlarged_image_index = self.folder_images.index(fpath)
            self._display_enlarged_gallery_image()
        else:
            # Re-render to highlight selection
            self._render_gallery_grid()

    def _change_gallery_page(self, delta):
        self.current_page += delta
        self._render_gallery_grid()

    # ================= ENLARGED IN-GALLERY INSPECTION =================
    def _display_enlarged_gallery_image(self):
        if not self.folder_images or self.enlarged_image_index < 0:
            return

        idx = self.enlarged_image_index
        fpath = self.folder_images[idx]
        total = len(self.folder_images)

        self.enl_counter_lbl.setText(f"Image {idx + 1} of {total}: {fpath.name}")
        self.btn_enl_prev.setEnabled(idx > 0)
        self.btn_enl_next.setEnabled(idx < total - 1)

        pm = QPixmap(str(fpath))
        if not pm.isNull():
            vw = max(100, self.enl_scroll.viewport().width() - 10)
            vh = max(100, self.enl_scroll.viewport().height() - 10)
            scaled = pm.scaled(vw, vh, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.enl_img_lbl.setPixmap(scaled)

            stat = fpath.stat()
            size_str = format_size(stat.st_size)
            dims = f"{pm.width()} × {pm.height()} px"
            self.enl_badge_lbl.setText(f"{fpath.suffix.upper().lstrip('.')} IMAGE")
            self.enl_specs_lbl.setText(f"{fpath.name}  •  {dims}  •  {size_str}")

        self.stack.setCurrentIndex(3)

    def _navigate_enlarged_image(self, delta):
        new_idx = self.enlarged_image_index + delta
        if 0 <= new_idx < len(self.folder_images):
            self.enlarged_image_index = new_idx
            self._display_enlarged_gallery_image()

    def _return_to_gallery_grid(self):
        self.stack.setCurrentIndex(2)
        self._render_gallery_grid()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.current_mode == "single" and self.single_view_stack.currentIndex() == 0 and self._fit_mode:
            self._update_single_image_display()
        elif self.stack.currentIndex() == 3 and self.enlarged_image_index >= 0:
            self._display_enlarged_gallery_image()

    # ================= THEME APPLICATION =================
    def apply_theme(self, theme="dark"):
        self.current_theme = theme
        super().apply_theme(theme)

        is_light = (theme == "light")
        bar_bg = "#f8fafc" if is_light else "#0e121d"
        border_col = "#cbd5e1" if is_light else "#232b43"
        txt_col = "#0f172a" if is_light else "#f1f5f9"
        sub_col = "#64748b" if is_light else "#8e98b0"
        btn_bg = "#ffffff" if is_light else "#1e243b"
        btn_border = "#cbd5e1" if is_light else "#2e3654"

        frame_style = f"background-color: {bar_bg}; border: 1px solid {border_col}; border-radius: 8px;"
        for f in (self.single_top_bar, self.single_specs_frame, self.gal_header_frame, self.gal_nav_frame, self.enl_top_bar, self.enl_specs_frame):
            f.setStyleSheet(frame_style)

        btn_style = f"""
            QPushButton {{
                background-color: {btn_bg};
                border: 1px solid {btn_border};
                border-radius: 6px;
                color: {txt_col};
                padding: 4px 10px;
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {'#eff6ff' if is_light else '#28304e'};
                border-color: #3b82f6;
            }}
            QPushButton:disabled {{
                background-color: {'#f1f5f9' if is_light else '#161b2b'};
                color: {'#94a3b8' if is_light else '#475569'};
                border-color: {btn_border};
            }}
        """
        for b in (self.btn_fit, self.btn_zoom_out, self.btn_zoom_in, self.btn_open_external, self.btn_gal_prev, self.btn_gal_next, self.btn_back_to_grid, self.btn_enl_prev, self.btn_enl_next):
            b.setStyleSheet(btn_style)

        if self.stack.currentIndex() == 2:
            self._render_gallery_grid()


class SmallGalleryView(QFrame):
    """
    Compact evidence preview gallery widget positioned directly beneath the drop box.
    Provides instant visual feedback for selected evidence:
    - Image: Scaled thumbnail preview (max height ~110px, width ~120px), file name, resolution, file size, format badge, and '👁️ Full View' button.
    - Video: Video icon, format badge, size, duration & resolution if parsed, and '▶ Open Video' button.
    - Audio: Audio icon, format badge, size, duration & bitrate if parsed, and '▶ Play Audio' button.
    - Document: Document icon, format badge, size, pages if parsed, and '📄 Open Document' button.
    - Folder: Folder icon, folder name, item count, and '📂 Open Folder' button.
    - Empty: Clean subtle placeholder encouraging drag & drop or file selection.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("SmallGalleryView")
        self.setFixedHeight(110)
        self.current_path = None
        self.current_record = None
        self._orig_pixmap = None
        self._theme = "dark"
        self._init_ui()
        self.apply_theme("dark")

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(4)

        self.stack = QStackedWidget()

        # Page 0: Empty Placeholder
        self.empty_page = QWidget()
        emp_layout = QVBoxLayout(self.empty_page)
        emp_layout.setContentsMargins(0, 0, 0, 0)
        emp_layout.setAlignment(Qt.AlignCenter)
        self.lbl_empty_icon = QLabel()
        self.lbl_empty_icon.setAlignment(Qt.AlignCenter)
        setup_label_icon(self.lbl_empty_icon, "eye", size=22, default_color="#3b82f6")
        self.lbl_empty_msg = QLabel("No file selected\nDrop a file or browse above to preview")
        self.lbl_empty_msg.setAlignment(Qt.AlignCenter)
        self.lbl_empty_msg.setStyleSheet("font-size: 11px; font-weight: 600; background: transparent; border: none;")
        emp_layout.addWidget(self.lbl_empty_icon)
        emp_layout.addWidget(self.lbl_empty_msg)
        self.stack.addWidget(self.empty_page)

        # Page 1: Image Thumbnail Preview
        self.image_page = QWidget()
        img_layout = QHBoxLayout(self.image_page)
        img_layout.setContentsMargins(4, 2, 4, 2)
        img_layout.setSpacing(10)

        self.thumb_container = QFrame()
        self.thumb_container.setFixedSize(95, 95)
        thumb_layout = QVBoxLayout(self.thumb_container)
        thumb_layout.setContentsMargins(2, 2, 2, 2)
        thumb_layout.setAlignment(Qt.AlignCenter)
        self.img_thumb_lbl = QLabel()
        self.img_thumb_lbl.setAlignment(Qt.AlignCenter)
        self.img_thumb_lbl.setCursor(QCursor(Qt.PointingHandCursor))
        self.img_thumb_lbl.mousePressEvent = lambda e: self._open_in_system_viewer()
        thumb_layout.addWidget(self.img_thumb_lbl)
        img_layout.addWidget(self.thumb_container)

        info_layout = QVBoxLayout()
        info_layout.setContentsMargins(0, 2, 0, 2)
        info_layout.setSpacing(3)

        self.lbl_img_title = QLabel("file.png")
        self.lbl_img_title.setStyleSheet("font-size: 12px; font-weight: 700; background: transparent; border: none;")
        self.lbl_img_title.setWordWrap(True)

        self.lbl_img_badge = QLabel("PNG IMAGE")
        self.lbl_img_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #38bdf8; background: rgba(56, 189, 248, 0.15); border-radius: 4px; padding: 2px 6px;")
        self.lbl_img_badge.setFixedHeight(18)

        badge_row = QHBoxLayout()
        badge_row.setSpacing(6)
        badge_row.addWidget(self.lbl_img_badge)
        badge_row.addStretch()

        self.lbl_img_specs = QLabel("1920 × 1080 px • 1.2 MB")
        self.lbl_img_specs.setStyleSheet("font-size: 11px; background: transparent; border: none;")

        info_layout.addWidget(self.lbl_img_title)
        info_layout.addLayout(badge_row)
        info_layout.addWidget(self.lbl_img_specs)
        info_layout.addStretch()

        self.btn_open_img = QPushButton("View Image")
        self.btn_open_img.setIcon(LucideIcons.icon("eye", color="#2563eb", size=16))
        self.btn_open_img.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_open_img.setFixedHeight(28)
        self.btn_open_img.setFixedWidth(130)
        self.btn_open_img.clicked.connect(self._open_in_system_viewer)
        info_layout.addWidget(self.btn_open_img)

        img_layout.addLayout(info_layout, 1)
        self.stack.addWidget(self.image_page)

        # Page 2: Non-Image Slate
        self.slate_page = QWidget()
        slate_layout = QHBoxLayout(self.slate_page)
        slate_layout.setContentsMargins(4, 2, 4, 2)
        slate_layout.setSpacing(10)

        self.slate_icon_box = QFrame()
        self.slate_icon_box.setFixedSize(76, 120)
        box_layout = QVBoxLayout(self.slate_icon_box)
        box_layout.setContentsMargins(0, 0, 0, 0)
        box_layout.setAlignment(Qt.AlignCenter)
        self.lbl_slate_icon = QLabel()
        self.lbl_slate_icon.setAlignment(Qt.AlignCenter)
        self.lbl_slate_icon.setStyleSheet("font-size: 34px; background: transparent; border: none;")
        box_layout.addWidget(self.lbl_slate_icon)
        slate_layout.addWidget(self.slate_icon_box)

        slate_info_layout = QVBoxLayout()
        slate_info_layout.setContentsMargins(0, 2, 0, 2)
        slate_info_layout.setSpacing(3)

        self.lbl_slate_title = QLabel("video.mp4")
        self.lbl_slate_title.setStyleSheet("font-size: 12px; font-weight: 700; background: transparent; border: none;")
        self.lbl_slate_title.setWordWrap(True)

        self.lbl_slate_badge = QLabel("MP4 VIDEO")
        self.lbl_slate_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #a78bfa; background: rgba(167, 139, 250, 0.15); border-radius: 4px; padding: 2px 6px;")
        self.lbl_slate_badge.setFixedHeight(18)

        slate_badge_row = QHBoxLayout()
        slate_badge_row.setSpacing(6)
        slate_badge_row.addWidget(self.lbl_slate_badge)
        slate_badge_row.addStretch()

        self.lbl_slate_desc = QLabel("Video File")
        self.lbl_slate_desc.setStyleSheet("font-size: 11px; background: transparent; border: none;")

        slate_info_layout.addWidget(self.lbl_slate_title)
        slate_info_layout.addLayout(slate_badge_row)
        slate_info_layout.addWidget(self.lbl_slate_desc)
        slate_info_layout.addStretch()

        self.btn_slate_action = QPushButton("▶ Open Evidence")
        self.btn_slate_action.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_slate_action.setFixedHeight(28)
        self.btn_slate_action.setFixedWidth(140)
        self.btn_slate_action.clicked.connect(self._open_in_system_viewer)
        slate_info_layout.addWidget(self.btn_slate_action)

        slate_layout.addLayout(slate_info_layout, 1)
        self.stack.addWidget(self.slate_page)

        layout.addWidget(self.stack)
        self.stack.setCurrentIndex(0)

    def load_file(self, file_path, record=None):
        if not file_path or not os.path.exists(file_path):
            self.current_path = None
            self.current_record = None
            self._orig_pixmap = None
            self.stack.setCurrentIndex(0)
            return

        p = Path(file_path)
        self.current_path = str(p)
        self.current_record = record
        ext = p.suffix.lower()

        image_exts = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif", ".tiff", ".tif", ".ico", ".svg", ".dng", ".cr2", ".nef", ".arw"}
        video_exts = {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm", ".m4v", ".flv"}
        audio_exts = {".mp3", ".wav", ".aac", ".flac", ".ogg", ".m4a", ".wma"}
        doc_exts = {".pdf", ".docx", ".doc", ".xlsx", ".pptx", ".txt", ".csv", ".json", ".xml", ".html"}

        if p.is_dir():
            self._orig_pixmap = None
            self.stack.setCurrentIndex(2)
            setup_label_icon(self.lbl_slate_icon, "folder", size=28, default_color=ic_col)
            self.lbl_slate_title.setText(p.name)
            self.lbl_slate_title.setToolTip(str(p.resolve()))
            self.lbl_slate_badge.setText("FOLDER")
            self.lbl_slate_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #38bdf8; background: rgba(56, 189, 248, 0.15); border-radius: 4px; padding: 2px 6px;")
            try:
                item_count = len(list(p.iterdir()))
                self.lbl_slate_desc.setText(f"{item_count} items in folder\nPath: {p.parent.name}/{p.name}")
            except Exception:
                self.lbl_slate_desc.setText(f"Folder\nPath: {p.parent.name}/{p.name}")
            self.btn_slate_action.setText("Open Folder")
            self.btn_slate_action.setIcon(LucideIcons.icon("folder", color=ic_col, size=16))
            return

        size_bytes = p.stat().st_size
        size_str = format_size(size_bytes)
        parsed = record.get("parsed", {}) if record else {}

        ic_col = "#2563eb" if getattr(self, "_theme", "dark") == "light" else "#60a5fa"

        if ext in image_exts:
            pm = QPixmap(str(p))
            if not pm.isNull():
                self._orig_pixmap = pm
                scaled = pm.scaled(112, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.img_thumb_lbl.setPixmap(scaled)
                self.lbl_img_title.setText(p.name)
                self.lbl_img_title.setToolTip(f"{p.name}\nSize: {size_str} ({size_bytes:,} bytes)\nPath: {p.resolve()}")
                self.lbl_img_badge.setText(f"{ext.upper().lstrip('.')} IMAGE")
                self.lbl_img_specs.setText(f"{pm.width()} × {pm.height()} px • {size_str}")
                self.btn_open_img.setText("View Image")
                self.btn_open_img.setIcon(LucideIcons.icon("eye", color=ic_col, size=16))
                self.stack.setCurrentIndex(1)
                return

        self._orig_pixmap = None
        self.stack.setCurrentIndex(2)
        self.lbl_slate_title.setText(p.name)
        self.lbl_slate_title.setToolTip(f"{p.name}\nSize: {size_str}\nPath: {p.resolve()}")

        if ext in video_exts:
            setup_label_icon(self.lbl_slate_icon, "play", size=28, default_color=ic_col)
            self.lbl_slate_badge.setText(f"{ext.upper().lstrip('.')} VIDEO")
            self.lbl_slate_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #a78bfa; background: rgba(167, 139, 250, 0.15); border-radius: 4px; padding: 2px 6px;")
            duration = parsed.get("Duration", parsed.get("MediaDuration", ""))
            resolution = parsed.get("ImageSize", "")
            specs = [f"Size: {size_str}"]
            if duration:
                specs.append(f"Duration: {duration}")
            if resolution:
                specs.append(f"Res: {resolution}")
            self.lbl_slate_desc.setText(" • ".join(specs))
            self.btn_slate_action.setText("Play Video")
            self.btn_slate_action.setIcon(LucideIcons.icon("play", color=ic_col, size=16))

        elif ext in audio_exts:
            setup_label_icon(self.lbl_slate_icon, "play", size=28, default_color=ic_col)
            self.lbl_slate_badge.setText(f"{ext.upper().lstrip('.')} AUDIO")
            self.lbl_slate_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #f472b6; background: rgba(244, 114, 182, 0.15); border-radius: 4px; padding: 2px 6px;")
            duration = parsed.get("Duration", "")
            specs = [f"Size: {size_str}"]
            if duration:
                specs.append(f"Duration: {duration}")
            self.lbl_slate_desc.setText(" • ".join(specs))
            self.btn_slate_action.setText("Play Audio")
            self.btn_slate_action.setIcon(LucideIcons.icon("play", color=ic_col, size=16))

        elif ext in doc_exts:
            setup_label_icon(self.lbl_slate_icon, "file_text", size=28, default_color=ic_col)
            self.lbl_slate_badge.setText(f"{ext.upper().lstrip('.')} DOC")
            self.lbl_slate_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #38bdf8; background: rgba(56, 189, 248, 0.15); border-radius: 4px; padding: 2px 6px;")
            pages = parsed.get("PageCount", parsed.get("Pages", ""))
            specs = [f"Size: {size_str}"]
            if pages:
                specs.append(f"Pages: {pages}")
            self.lbl_slate_desc.setText(" • ".join(specs))
            self.btn_slate_action.setText("Open Document")
            self.btn_slate_action.setIcon(LucideIcons.icon("file_text", color=ic_col, size=16))

        else:
            setup_label_icon(self.lbl_slate_icon, "folder", size=28, default_color=ic_col)
            self.lbl_slate_badge.setText(f"{ext.upper().lstrip('.') or 'FILE'}")
            self.lbl_slate_badge.setStyleSheet("font-size: 10px; font-weight: 700; color: #94a3b8; background: rgba(148, 163, 184, 0.15); border-radius: 4px; padding: 2px 6px;")
            self.lbl_slate_desc.setText(f"File • {size_str}")
            self.btn_slate_action.setText("Open File")
            self.btn_slate_action.setIcon(LucideIcons.icon("folder", color=ic_col, size=16))

    def _open_in_system_viewer(self):
        if self.current_path and os.path.exists(self.current_path):
            open_evidence_file(self, self.current_path)

    def apply_theme(self, theme="dark"):
        self._theme = theme
        if theme == "light":
            self.setStyleSheet("#SmallGalleryView { background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; }")
            self.lbl_empty_icon.setStyleSheet("font-size: 26px; color: #94a3b8; background: transparent; border: none;")
            self.lbl_empty_msg.setStyleSheet("font-size: 11px; color: #64748b; font-weight: 600; background: transparent; border: none;")
            self.thumb_container.setStyleSheet("background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px;")
            self.lbl_img_title.setStyleSheet("font-size: 12px; font-weight: 700; color: #0f172a; background: transparent; border: none;")
            self.lbl_img_specs.setStyleSheet("font-size: 11px; color: #64748b; background: transparent; border: none;")
            self.btn_open_img.setStyleSheet("""
                QPushButton {
                    background-color: #f1f5f9;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    color: #0f172a;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover { background-color: #e2e8f0; border-color: #2563eb; }
            """)
            self.slate_icon_box.setStyleSheet("background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px;")
            self.lbl_slate_title.setStyleSheet("font-size: 12px; font-weight: 700; color: #0f172a; background: transparent; border: none;")
            self.lbl_slate_desc.setStyleSheet("font-size: 11px; color: #64748b; background: transparent; border: none;")
            self.btn_slate_action.setStyleSheet("""
                QPushButton {
                    background-color: #f1f5f9;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    color: #0f172a;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover { background-color: #e2e8f0; border-color: #2563eb; }
            """)
        else:
            self.setStyleSheet("#SmallGalleryView { background-color: #0e121d; border: 1px solid #232b43; border-radius: 8px; }")
            self.lbl_empty_icon.setStyleSheet("font-size: 26px; color: #475569; background: transparent; border: none;")
            self.lbl_empty_msg.setStyleSheet("font-size: 11px; color: #8e98b0; font-weight: 600; background: transparent; border: none;")
            self.thumb_container.setStyleSheet("background-color: #080a10; border: 1px solid #1c2236; border-radius: 6px;")
            self.lbl_img_title.setStyleSheet("font-size: 12px; font-weight: 700; color: #f1f5f9; background: transparent; border: none;")
            self.lbl_img_specs.setStyleSheet("font-size: 11px; color: #8e98b0; background: transparent; border: none;")
            self.btn_open_img.setStyleSheet("""
                QPushButton {
                    background-color: #1e243b;
                    border: 1px solid #2e3654;
                    border-radius: 6px;
                    color: #f1f5f9;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
            """)
            self.slate_icon_box.setStyleSheet("background-color: #080a10; border: 1px solid #1c2236; border-radius: 6px;")
            self.lbl_slate_title.setStyleSheet("font-size: 12px; font-weight: 700; color: #f1f5f9; background: transparent; border: none;")
            self.lbl_slate_desc.setStyleSheet("font-size: 11px; color: #8e98b0; background: transparent; border: none;")
            self.btn_slate_action.setStyleSheet("""
                QPushButton {
                    background-color: #1e243b;
                    border: 1px solid #2e3654;
                    border-radius: 6px;
                    color: #f1f5f9;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
            """)


# ==============================================================================
# SECTION 7: MAIN APPLICATION WINDOW & CONTROLLER
# ==============================================================================

def ensure_ui_assets():
    """Generate high-contrast, crystal-clear checkbox & radio button indicator assets,
    plus standard modern enterprise UI icons."""
    assets_dir = APP_ROOT / "assets"
    assets_dir.mkdir(exist_ok=True)

    cb_on_path = assets_dir / "cb_checked.png"
    cb_off_path = assets_dir / "cb_unchecked.png"
    cb_hover_path = assets_dir / "cb_hover.png"
    rb_on_path = assets_dir / "rb_checked.png"
    rb_off_path = assets_dir / "rb_unchecked.png"
    rb_hover_path = assets_dir / "rb_hover.png"

    cb_on_light_path = assets_dir / "cb_checked_light.png"
    cb_off_light_path = assets_dir / "cb_unchecked_light.png"
    cb_hover_light_path = assets_dir / "cb_hover_light.png"
    rb_on_light_path = assets_dir / "rb_checked_light.png"
    rb_off_light_path = assets_dir / "rb_unchecked_light.png"
    rb_hover_light_path = assets_dir / "rb_hover_light.png"

    if not (cb_on_path.exists() and cb_off_path.exists() and rb_on_path.exists() and rb_off_path.exists()):
        # Light Checkbox Unchecked: Crisp white rect with slate-400 border
        pm = QPixmap(24, 24); pm.fill(Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(255, 255, 255)); p.setPen(QPen(QColor(100, 116, 139), 2.0)); p.drawRoundedRect(2, 2, 20, 20, 5, 5); p.end()
        pm.save(str(cb_off_light_path))

        # Light Checkbox Hover: Crisp white rect with bright blue border
        pm = QPixmap(24, 24); pm.fill(Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(255, 255, 255)); p.setPen(QPen(QColor(37, 99, 235), 2.0)); p.drawRoundedRect(2, 2, 20, 20, 5, 5); p.end()
        pm.save(str(cb_hover_light_path))

        # Light Checkbox Checked: Vibrant blue rect with crisp white checkmark
        pm = QPixmap(24, 24); pm.fill(Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(37, 99, 235)); p.setPen(QPen(QColor(29, 78, 216), 1.5)); p.drawRoundedRect(1, 1, 21, 21, 5, 5)
        p.setPen(QPen(QColor(255, 255, 255), 2.6, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(6, 12, 10, 16); p.drawLine(10, 16, 17, 7); p.end()
        pm.save(str(cb_on_light_path))

        # Light Radio Unchecked: Crisp white circle with slate-400 border
        pm = QPixmap(24, 24); pm.fill(Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(255, 255, 255)); p.setPen(QPen(QColor(100, 116, 139), 2.0)); p.drawEllipse(2, 2, 20, 20); p.end()
        pm.save(str(rb_off_light_path))

        # Light Radio Hover: Crisp white circle with bright blue border
        pm = QPixmap(24, 24); pm.fill(Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(255, 255, 255)); p.setPen(QPen(QColor(37, 99, 235), 2.0)); p.drawEllipse(2, 2, 20, 20); p.end()
        pm.save(str(rb_hover_light_path))

        # Light Radio Checked: Crisp white circle with bright blue ring and core dot
        pm = QPixmap(24, 24); pm.fill(Qt.transparent); p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(255, 255, 255)); p.setPen(QPen(QColor(37, 99, 235), 2.2)); p.drawEllipse(2, 2, 20, 20)
        p.setBrush(QColor(37, 99, 235)); p.setPen(Qt.NoPen); p.drawEllipse(7, 7, 10, 10); p.end()
        pm.save(str(rb_on_light_path))

    # 1. Checkbox Checked: Vibrant blue rect with crisp bold white checkmark
        pm = QPixmap(24, 24)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(37, 99, 235))
        p.setPen(QPen(QColor(96, 165, 250), 1.5))
        p.drawRoundedRect(1, 1, 21, 21, 5, 5)
        pen = QPen(QColor(255, 255, 255), 2.6)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        p.setPen(pen)
        p.drawLine(6, 12, 10, 16)
        p.drawLine(10, 16, 17, 7)
        p.end()
        pm.save(str(cb_on_path))

        # 2. Checkbox Unchecked: Dark slate rect with high-contrast slate-gray border
        pm = QPixmap(24, 24)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(14, 18, 29))
        p.setPen(QPen(QColor(100, 116, 139), 2.0))
        p.drawRoundedRect(2, 2, 20, 20, 5, 5)
        p.end()
        pm.save(str(cb_off_path))

        # 3. Checkbox Hover: Dark rect with bright blue border
        pm = QPixmap(24, 24)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(20, 26, 42))
        p.setPen(QPen(QColor(96, 165, 250), 2.0))
        p.drawRoundedRect(2, 2, 20, 20, 5, 5)
        p.end()
        pm.save(str(cb_hover_path))

        # 4. Radio Checked: Dark circle with bright blue border and distinct white/blue core
        pm = QPixmap(24, 24)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(14, 18, 29))
        p.setPen(QPen(QColor(59, 130, 246), 2.0))
        p.drawEllipse(2, 2, 20, 20)
        p.setBrush(QColor(37, 99, 235))
        p.setPen(Qt.NoPen)
        p.drawEllipse(6, 6, 12, 12)
        p.setBrush(QColor(255, 255, 255))
        p.drawEllipse(9, 9, 6, 6)
        p.end()
        pm.save(str(rb_on_path))

        # 5. Radio Unchecked: Dark circle with clear slate-gray border
        pm = QPixmap(24, 24)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(14, 18, 29))
        p.setPen(QPen(QColor(100, 116, 139), 2.0))
        p.drawEllipse(2, 2, 20, 20)
        p.end()
        pm.save(str(rb_off_path))

        # 6. Radio Hover: Dark circle with bright blue border
        pm = QPixmap(24, 24)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(20, 26, 42))
        p.setPen(QPen(QColor(96, 165, 250), 2.0))
        p.drawEllipse(2, 2, 20, 20)
        p.end()
        pm.save(str(rb_hover_path))

    # Standard Modern Vector UI Icons
    def _draw_canvas(w=32, h=32):
        pm = QPixmap(w, h)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        return pm, p

    # 1. Compare Icon (Dual documents with bidirectional comparison arrows)
    cmp_path = assets_dir / "compare.png"
    if not cmp_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(59, 130, 246), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(3, 4, 11, 15, 2, 2)
        p.setPen(QPen(QColor(148, 163, 184), 1.2))
        p.drawLine(6, 8, 11, 8); p.drawLine(6, 11, 11, 11); p.drawLine(6, 14, 9, 14)
        p.setPen(QPen(QColor(34, 211, 238), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(18, 4, 11, 15, 2, 2)
        p.setPen(QPen(QColor(148, 163, 184), 1.2))
        p.drawLine(21, 8, 26, 8); p.drawLine(21, 11, 26, 11); p.drawLine(21, 14, 24, 14)
        p.setPen(QPen(QColor(241, 245, 249), 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(7, 24, 25, 24); p.drawLine(22, 21, 25, 24); p.drawLine(22, 27, 24, 25)
        p.drawLine(7, 28, 25, 28); p.drawLine(10, 25, 7, 28); p.drawLine(10, 31, 7, 28)
        p.end(); pm.save(str(cmp_path))

    # 2. Shield Icon
    sh_path = assets_dir / "shield.png"
    if not sh_path.exists():
        pm, p = _draw_canvas(32, 32)
        path = QPainterPath(); path.moveTo(16, 3); path.lineTo(27, 7); path.quadTo(27, 21, 16, 29); path.quadTo(5, 21, 5, 7); path.closeSubpath()
        p.setPen(QPen(QColor(96, 165, 250), 1.8)); p.setBrush(QColor(22, 28, 48)); p.drawPath(path)
        p.setPen(QPen(QColor(52, 211, 153), 2.0, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(11, 15, 14, 19); p.drawLine(14, 19, 21, 11)
        p.end(); pm.save(str(sh_path))

    # 3. Checksum
    chk_path = assets_dir / "checksum.png"
    if not chk_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(59, 130, 246), 1.6)); p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(4, 4, 24, 24, 5, 5)
        p.setPen(QPen(QColor(96, 165, 250), 2.2, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(12, 9, 10, 23); p.drawLine(20, 9, 18, 23)
        p.drawLine(8, 13, 23, 13); p.drawLine(7, 19, 22, 19)
        p.end(); pm.save(str(chk_path))

    # 4. Document
    doc_path = assets_dir / "document.png"
    if not doc_path.exists():
        pm, p = _draw_canvas(32, 32)
        dp = QPainterPath(); dp.moveTo(7, 4); dp.lineTo(19, 4); dp.lineTo(25, 10); dp.lineTo(25, 28); dp.lineTo(7, 28); dp.closeSubpath()
        p.setPen(QPen(QColor(148, 163, 184), 1.8)); p.setBrush(QColor(18, 24, 40)); p.drawPath(dp)
        p.drawLine(19, 4, 19, 10); p.drawLine(19, 10, 25, 10)
        p.setPen(QPen(QColor(100, 116, 139), 1.6))
        p.drawLine(11, 14, 21, 14); p.drawLine(11, 18, 21, 18); p.drawLine(11, 22, 17, 22)
        p.end(); pm.save(str(doc_path))

    # 5. Filter
    flt_path = assets_dir / "filter.png"
    if not flt_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(148, 163, 184), 1.8, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(6, 9, 26, 9); p.drawLine(6, 16, 26, 16); p.drawLine(6, 23, 26, 23)
        p.setBrush(QColor(59, 130, 246)); p.setPen(QPen(QColor(96, 165, 250), 1.2))
        p.drawEllipse(11, 6, 6, 6); p.drawEllipse(20, 13, 6, 6); p.drawEllipse(14, 20, 6, 6)
        p.end(); pm.save(str(flt_path))

    # 6. Columns
    col_path = assets_dir / "columns.png"
    if not col_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(148, 163, 184), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(4, 5, 24, 22, 3, 3)
        p.drawLine(12, 5, 12, 27); p.drawLine(20, 5, 20, 27); p.drawLine(4, 11, 28, 11)
        p.end(); pm.save(str(col_path))

    # 7. Settings
    set_path = assets_dir / "settings.png"
    if not set_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(148, 163, 184), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawEllipse(8, 8, 16, 16)
        p.setPen(QPen(QColor(148, 163, 184), 2.4, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(16, 3, 16, 7); p.drawLine(16, 25, 16, 29)
        p.drawLine(3, 16, 7, 16); p.drawLine(25, 16, 29, 16)
        p.drawLine(7, 7, 10, 10); p.drawLine(22, 22, 25, 25)
        p.drawLine(25, 7, 22, 10); p.drawLine(10, 22, 7, 25)
        p.setBrush(QColor(11, 14, 23)); p.setPen(QPen(QColor(59, 130, 246), 1.5))
        p.drawEllipse(13, 13, 6, 6)
        p.end(); pm.save(str(set_path))

    # 8. Export
    exp_path = assets_dir / "export.png"
    if not exp_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(148, 163, 184), 1.8)); p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(6, 12, 20, 15, 3, 3)
        p.setPen(QPen(QColor(52, 211, 153), 2.2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(16, 5, 16, 18); p.drawLine(11, 10, 16, 5); p.drawLine(21, 10, 16, 5)
        p.end(); pm.save(str(exp_path))

    # 9. Storage
    stg_path = assets_dir / "storage.png"
    if not stg_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(96, 165, 250), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(5, 7, 22, 18, 3, 3)
        p.setPen(QPen(QColor(100, 116, 139), 1.5))
        p.drawLine(5, 19, 27, 19)
        p.setBrush(QColor(52, 211, 153)); p.setPen(Qt.NoPen)
        p.drawEllipse(22, 21, 2, 2)
        p.end(); pm.save(str(stg_path))

    # 10. Timer
    tm_path = assets_dir / "timer.png"
    if not tm_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(96, 165, 250), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawEllipse(6, 7, 20, 20)
        p.setPen(QPen(QColor(96, 165, 250), 2.0, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(16, 3, 16, 6)
        p.drawLine(16, 17, 16, 11); p.drawLine(16, 17, 21, 17)
        p.end(); pm.save(str(tm_path))

    # 11. Trash
    tr_path = assets_dir / "trash.png"
    if not tr_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(239, 68, 68), 1.8, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(9, 10, 14, 17, 2, 2)
        p.drawLine(6, 10, 26, 10)
        p.drawLine(13, 6, 19, 6)
        p.drawLine(13, 14, 13, 23); p.drawLine(19, 14, 19, 23)
        p.end(); pm.save(str(tr_path))

    # 12. Paste
    pst_path = assets_dir / "paste.png"
    if not pst_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(148, 163, 184), 1.8)); p.setBrush(QColor(18, 24, 40))
        p.drawRoundedRect(7, 7, 18, 21, 2, 2)
        p.setBrush(QColor(59, 130, 246))
        p.drawRoundedRect(11, 4, 10, 5, 2, 2)
        p.setPen(QPen(QColor(100, 116, 139), 1.5))
        p.drawLine(11, 15, 21, 15); p.drawLine(11, 19, 21, 19); p.drawLine(11, 23, 17, 23)
        p.end(); pm.save(str(pst_path))

    # 13. Excel
    xl_path = assets_dir / "excel.png"
    if not xl_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(16, 185, 129), 1.8)); p.setBrush(QColor(16, 185, 129, 35))
        p.drawRoundedRect(4, 4, 24, 24, 4, 4)
        p.setPen(QPen(QColor(52, 211, 153), 2.4, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(11, 10, 21, 22); p.drawLine(21, 10, 11, 22)
        p.end(); pm.save(str(xl_path))

    # 14. HTML
    htm_path = assets_dir / "html.png"
    if not htm_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(59, 130, 246), 1.8)); p.setBrush(QColor(59, 130, 246, 35))
        p.drawRoundedRect(4, 4, 24, 24, 4, 4)
        p.setPen(QPen(QColor(96, 165, 250), 2.2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(12, 12, 8, 16); p.drawLine(8, 16, 12, 20)
        p.drawLine(20, 12, 24, 16); p.drawLine(24, 16, 20, 20)
        p.drawLine(18, 10, 14, 22)
        p.end(); pm.save(str(htm_path))

    # 15. CSV
    csv_path = assets_dir / "csv.png"
    if not csv_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(245, 158, 11), 1.8)); p.setBrush(QColor(245, 158, 11, 35))
        p.drawRoundedRect(4, 4, 24, 24, 4, 4)
        p.setPen(QPen(QColor(251, 191, 36), 1.6))
        p.drawLine(7, 12, 25, 12); p.drawLine(7, 20, 25, 20)
        p.drawLine(13, 7, 13, 25); p.drawLine(19, 7, 19, 25)
        p.end(); pm.save(str(csv_path))

    # 16. Check badge
    chk_b_path = assets_dir / "check_badge.png"
    if not chk_b_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setBrush(QColor(16, 185, 129)); p.setPen(Qt.NoPen)
        p.drawEllipse(4, 4, 24, 24)
        p.setPen(QPen(QColor(255, 255, 255), 2.6, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.drawLine(10, 16, 14, 20); p.drawLine(14, 20, 22, 11)
        p.end(); pm.save(str(chk_b_path))

    # 17. Cross badge
    crs_b_path = assets_dir / "cross_badge.png"
    if not crs_b_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setBrush(QColor(239, 68, 68)); p.setPen(Qt.NoPen)
        p.drawEllipse(4, 4, 24, 24)
        p.setPen(QPen(QColor(255, 255, 255), 2.6, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(11, 11, 21, 21); p.drawLine(21, 11, 11, 21)
        p.end(); pm.save(str(crs_b_path))

    # 18. Info badge
    inf_b_path = assets_dir / "info_badge.png"
    if not inf_b_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setBrush(QColor(59, 130, 246)); p.setPen(Qt.NoPen)
        p.drawEllipse(4, 4, 24, 24)
        p.setPen(QPen(QColor(255, 255, 255), 2.4, Qt.SolidLine, Qt.RoundCap))
        p.drawPoint(16, 10); p.drawLine(16, 14, 16, 21)
        p.end(); pm.save(str(inf_b_path))

    # 19. Activity
    act_path = assets_dir / "activity.png"
    if not act_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(59, 130, 246), 2.0, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.setBrush(Qt.NoBrush)
        p.drawLine(3, 16, 9, 16); p.drawLine(9, 16, 13, 8); p.drawLine(13, 8, 17, 24); p.drawLine(17, 24, 21, 12); p.drawLine(21, 12, 24, 16); p.drawLine(24, 16, 29, 16)
        p.end(); pm.save(str(act_path))

    # 20. Inspect
    ins_path = assets_dir / "inspect.png"
    if not ins_path.exists():
        pm, p = _draw_canvas(32, 32)
        p.setPen(QPen(QColor(96, 165, 250), 2.2)); p.setBrush(QColor(18, 24, 40))
        p.drawEllipse(6, 6, 14, 14)
        p.setPen(QPen(QColor(96, 165, 250), 2.6, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(17, 17, 26, 26)
        p.end(); pm.save(str(ins_path))

    # Central desktop branding; HTML reports use separate legacy artwork.
    root_dir = APP_ROOT
    logo_path = root_dir / "resources" / "branding" / "favicon.png"

    process_src = root_dir / "process.png"
    clean_proc_path = assets_dir / "process_clean.png"
    if process_src.exists() and not clean_proc_path.exists():
        try:
            from PIL import Image
            import numpy as np
            im = Image.open(str(process_src)).convert('RGBA')
            arr = np.array(im, dtype=np.float32)
            dark = np.maximum(np.maximum(arr[:,:,0], arr[:,:,1]), arr[:,:,2])
            alpha = np.clip((dark - 10) / 35.0, 0.0, 1.0) * 255.0
            arr[:, :, 3] = alpha
            clean_im = Image.fromarray(arr.astype(np.uint8), mode='RGBA')
            clean_im.save(str(clean_proc_path))
        except Exception:
            pass

    proc_path = clean_proc_path if clean_proc_path.exists() else (process_src if process_src.exists() else (assets_dir / "activity.png"))

    return {
        "cb_on": str(cb_on_path).replace("\\", "/"),
        "cb_off": str(cb_off_path).replace("\\", "/"),
        "cb_hover": str(cb_hover_path).replace("\\", "/"),
        "rb_on": str(rb_on_path).replace("\\", "/"),
        "rb_off": str(rb_off_path).replace("\\", "/"),
        "rb_hover": str(rb_hover_path).replace("\\", "/"),
        "cb_on_light": str(assets_dir / "cb_checked_light.png").replace("\\", "/"),
        "cb_off_light": str(assets_dir / "cb_unchecked_light.png").replace("\\", "/"),
        "cb_hover_light": str(assets_dir / "cb_hover_light.png").replace("\\", "/"),
        "rb_on_light": str(assets_dir / "rb_checked_light.png").replace("\\", "/"),
        "rb_off_light": str(assets_dir / "rb_unchecked_light.png").replace("\\", "/"),
        "rb_hover_light": str(assets_dir / "rb_hover_light.png").replace("\\", "/"),
        "compare": str(assets_dir / "compare.png").replace("\\", "/"),
        "shield": str(assets_dir / "shield.png").replace("\\", "/"),
        "checksum": str(assets_dir / "checksum.png").replace("\\", "/"),
        "document": str(assets_dir / "document.png").replace("\\", "/"),
        "filter": str(assets_dir / "filter.png").replace("\\", "/"),
        "columns": str(assets_dir / "columns.png").replace("\\", "/"),
        "settings": str(assets_dir / "settings.png").replace("\\", "/"),
        "export": str(assets_dir / "export.png").replace("\\", "/"),
        "storage": str(assets_dir / "storage.png").replace("\\", "/"),
        "timer": str(assets_dir / "timer.png").replace("\\", "/"),
        "trash": str(assets_dir / "trash.png").replace("\\", "/"),
        "paste": str(assets_dir / "paste.png").replace("\\", "/"),
        "excel": str(assets_dir / "excel.png").replace("\\", "/"),
        "html": str(assets_dir / "html.png").replace("\\", "/"),
        "csv": str(assets_dir / "csv.png").replace("\\", "/"),
        "check_badge": str(assets_dir / "check_badge.png").replace("\\", "/"),
        "cross_badge": str(assets_dir / "cross_badge.png").replace("\\", "/"),
        "info_badge": str(assets_dir / "info_badge.png").replace("\\", "/"),
        "activity": str(assets_dir / "activity.png").replace("\\", "/"),
        "inspect": str(assets_dir / "inspect.png").replace("\\", "/"),
        "logo": str(logo_path).replace("\\", "/"),
        "process": str(proc_path).replace("\\", "/"),
        "process_raw": str(process_src if process_src.exists() else proc_path).replace("\\", "/"),
    }


class FileInventoryProApp(QMainWindow):
    """
    Main Application Window featuring Surface Modern Slate Dashboard styling,
    smooth responsiveness, cursor-responsive buttons, global Dark/Light theme switching,
    and three enterprise workstation modules:
    1. File Inventory & Audit
    2. Cryptographic Hash Verification Suite
    3. Metadata Explorer (Official Phil Harvey ExifTool Engine Integration)
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_WINDOW_TITLE)

        # Responsive adaptive sizing based on current screen geometry and aspect ratio
        screen = QApplication.primaryScreen()
        if screen:
            avail = screen.availableGeometry()
            init_w = min(1320, max(MIN_WINDOW_WIDTH, int(avail.width() * 0.88)))
            init_h = min(880, max(MIN_WINDOW_HEIGHT, int(avail.height() * 0.88)))
            self.resize(init_w, init_h)
        else:
            self.resize(1200, 780)

        self.setMinimumSize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)

        # State storage - Core & Existing Modules
        self.inventory_worker = None
        self.hash_worker = None
        self.hash_items = []  # List of dicts for verification mode
        self.last_generated_files = []
        self.start_time = None
        self.timer_elapsed = QTimer(self)
        self.timer_elapsed.timeout.connect(self._update_elapsed_display)

        # State storage - Settings persistence
        self.settings_store = QSettings("Antigravity", "FileInventoryPro")
        configure_hash_preferences(self.settings_store)
        saved_theme = self.settings_store.value("theme", "dark")
        self.current_theme = saved_theme if saved_theme in ("dark", "light") else "dark"
        install_message_box_theme_filter()
        self.custom_exiftool_path = self.settings_store.value("custom_exiftool_path", None)
        self.active_nav_index = 0
        self.active_mode = "overview"

        # State storage - Metadata Explorer Module
        self.metadata_worker = None
        self.active_meta_record = None
        self.all_meta_records = []

        # Session Activity Log (Real operations log for Overview screen)
        self.session_activity_log = []

        # Detect ExifTool engine dynamically at application startup
        self.exiftool_ver, self.exiftool_ok, self.exiftool_path = ExifToolEngine.get_version(self.custom_exiftool_path)

        # UI Assets & High-DPI Vector Icons
        self.assets = ensure_ui_assets()

        # Set Application & Window Icon
        if self.assets.get("logo") and os.path.exists(self.assets["logo"]):
            self.setWindowIcon(application_icon())

        # Build UI Structure first
        self._build_ui()

        # Restore startup tab (Overview by default)
        saved_tab = self.settings_store.value("startup_tab", 0, type=int)
        startup_index = saved_tab if saved_tab in (0, 1, 2, 3, 4) else 0
        self._set_active_nav_index(startup_index)

        # Apply Global Modern QSS Theme (persisted theme)
        self._apply_global_theme(self.current_theme)

        # Apply Default Preset
        self._apply_column_preset("media")

    def _style_button_primary(self, btn, theme="dark"):
        if not btn: return
        if theme == "light":
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #2563eb;
                    color: #ffffff;
                    font-weight: 700;
                    font-size: 12px;
                    border: none;
                    border-radius: 8px;
                    padding: 8px 16px;
                }
                QPushButton:hover { background-color: #1d4ed8; }
                QPushButton:disabled { background-color: #bfdbfe; color: #ffffff; }
            """)
        else:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #3b82f6;
                    color: #ffffff;
                    font-weight: 700;
                    font-size: 12px;
                    border: none;
                    border-radius: 8px;
                    padding: 8px 16px;
                }
                QPushButton:hover { background-color: #60a5fa; }
                QPushButton:disabled { background-color: #161b2b; color: #55607c; border: 1px solid #232b43; }
            """)

    def _style_button_secondary(self, btn, theme="dark"):
        if not btn: return
        if theme == "light":
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    border: 1px solid #cbd5e1;
                    border-radius: 8px;
                    color: #0f172a;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #f1f5f9;
                    border-color: #2563eb;
                    color: #1d4ed8;
                }
                QPushButton:disabled {
                    background-color: #f8fafc;
                    border-color: #e2e8f0;
                    color: #94a3b8;
                }
            """)
        else:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #1e243b;
                    border: 1px solid #2e3654;
                    border-radius: 8px;
                    color: #f1f5f9;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #28304e;
                    border-color: #3b82f6;
                    color: #ffffff;
                }
                QPushButton:disabled {
                    background-color: #121727;
                    border-color: #1e263d;
                    color: #55607c;
                }
            """)

    def _style_button_danger(self, btn, theme="dark"):
        if not btn: return
        if theme == "light":
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #ffffff;
                    border: 1px solid #fca5a5;
                    border-radius: 8px;
                    color: #dc2626;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #fef2f2;
                    border-color: #ef4444;
                    color: #b91c1c;
                }
                QPushButton:disabled {
                    background-color: #f1f5f9;
                    border-color: #e2e8f0;
                    color: #94a3b8;
                }
            """)
        else:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: rgba(239, 68, 68, 0.15);
                    border: 1px solid rgba(239, 68, 68, 0.35);
                    border-radius: 8px;
                    color: #f87171;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: rgba(239, 68, 68, 0.25);
                    border-color: #ef4444;
                    color: #ffffff;
                }
                QPushButton:disabled {
                    background-color: #121727;
                    border-color: #1e263d;
                    color: #55607c;
                }
            """)

    def _update_all_workstation_buttons(self, theme="dark"):
        """Synchronize every button across all 5 modules with current theme palette."""
        # 1. Primary Buttons
        primary_list = [
            getattr(self, "hash_btn_add_files", None),
            getattr(self, "hash_btn_set_target", None),
            getattr(self, "btn_verify_hashes", None),
            getattr(self, "btn_run_inventory", None),
            getattr(self, "meta_btn_extract", None),
            getattr(self, "settings_btn_save", None),
            getattr(self, "settings_btn_test_exif", None),
        ]
        if hasattr(self, "overview_shortcut_buttons"):
            primary_list.extend(self.overview_shortcut_buttons)

        for btn in primary_list:
            if btn: self._style_button_primary(btn, theme)

        # 2. Secondary Buttons
        secondary_list = [
            getattr(self, "hash_btn_add_folder", None),
            getattr(self, "hash_btn_load_checksum", None),
            getattr(self, "hash_btn_compare", None),
            getattr(self, "hash_btn_paste", None),
            getattr(self, "hash_btn_compare2", None),
            getattr(self, "hash_btn_clear", None),
            getattr(self, "hash_btn_clear_tgt", None),
            getattr(self, "inv_btn_src_browse", None),
            getattr(self, "inv_btn_out_browse", None),
            getattr(self, "meta_btn_browse", None),
            getattr(self, "meta_btn_browse_folder", None),
            getattr(self, "meta_btn_browse_out", None),
            getattr(self, "meta_btn_config_engine", None),
            getattr(self, "meta_btn_open_processed_folder", None),
            getattr(self, "meta_btn_expand", None),
            getattr(self, "meta_btn_collapse", None),
            getattr(self, "meta_btn_copy", None),
            getattr(self, "meta_btn_view_json", None),
            getattr(self, "header_compare_btn", None),
            getattr(self, "settings_btn_browse_exif", None),
            getattr(self, "settings_btn_auto_exif", None),
            getattr(self, "settings_btn_browse_out", None),
            getattr(self, "settings_btn_reset", None),
        ]
        if hasattr(self, "btn_qa_compare") and self.btn_qa_compare:
            ic_c = "#2563eb" if theme == "light" else "#60a5fa"
            self.btn_qa_compare.setIcon(LucideIcons.icon("compare", color=ic_c, size=16))
            self.btn_qa_hash.setIcon(LucideIcons.icon("shield", color=ic_c, size=16))
            self.btn_qa_inv.setIcon(LucideIcons.icon("folder", color=ic_c, size=16))
            self.btn_qa_meta.setIcon(LucideIcons.icon("search", color=ic_c, size=16))

        if hasattr(self, "hash_btn_clear") and self.hash_btn_clear:
            ic_c = "#2563eb" if theme == "light" else "#60a5fa"
            self.hash_btn_clear.setIcon(LucideIcons.icon("trash", color=ic_c, size=16))

        if hasattr(self, "overview_action_buttons"):
            secondary_list.extend(self.overview_action_buttons)
        if hasattr(self, "hash_export_buttons"):
            secondary_list.extend(self.hash_export_buttons)
        if hasattr(self, "meta_quick_html_btn"):
            secondary_list.extend([self.meta_quick_html_btn, self.meta_quick_txt_btn, self.meta_quick_json_btn])

        for btn in secondary_list:
            if btn: self._style_button_secondary(btn, theme)

        # Use the standard status-green palette for Excel export icons
        if hasattr(self, "exp_excel_btn") and self.exp_excel_btn:
            self.exp_excel_btn.setIcon(LucideIcons.icon("file_spreadsheet", color="#287a45" if theme == "light" else ("#80cc93").replace("#80cc93", status_green(self)), size=14))

        # 3. Danger / Cancel Buttons (Operational Cancels only)
        danger_list = [
            getattr(self, "btn_cancel_verify", None),
            getattr(self, "btn_cancel_inventory", None),
            getattr(self, "meta_btn_cancel", None),
        ]
        for btn in danger_list:
            if btn: self._style_button_danger(btn, theme)

        self._refresh_meta_tree_theme(theme)

    def _apply_global_theme(self, theme="dark"):
        """
        Inject clean Surface Dashboard Qt StyleSheet with standard, high-contrast indicators.
        Supports both Dark Forensic Workstation mode and Clean Light Workstation mode across
        all tabs and modules without resetting user inputs or state.
        """
        self.current_theme = theme
        self.settings_store.setValue("theme", theme)
        QApplication.instance().setProperty('findexTheme', theme)
        assets = ensure_ui_assets()

        if theme == "light":
            # 3-Color Rule (60-30-10) Classy Forensic Workstation Theme
            # 60% Dominant Canvas: Distinct cool workstation slate (#EBF0F7) clearly separating background from white cards
            bg_main = "#EBF0F7"
            # 30% Structural Surfaces & Hierarchy: Crisp pure white containers with clearly defined slate-300 borders
            bg_sidebar = "#FFFFFF"      # Crisp white modern sidebar
            bg_header = "#FFFFFF"       # Crisp white header
            bg_card = "#FFFFFF"         # Pure white elevated container surfaces (pops against #EBF0F7)
            bg_input = "#FFFFFF"        # Pure white input fields
            bg_table = "#FFFFFF"        # Pure white table background
            bg_head = "#F1F5F9"         # Clean distinct table header
            border_color = "#CBD5E1"    # Distinct Slate-300 card and separator borders (clearly separating containers from canvas)
            border_strong = "#94A3B8"   # Crisp Slate-400 border for controls
            border_subtle = "#E2E8F0"
            text_primary = "#0F172A"    # Deep dark slate (Slate-900) for maximum legibility
            text_secondary = "#475569"  # Balanced slate (Slate-600) for subtitles and labels
            # 10% Signature Accent: Focused forensic royal blue for active elements and icons
            accent_blue = "#2563EB"     # Royal Blue (Blue-600)
            accent_hover = "#1D4ED8"    # Deep Royal Blue (Blue-700)
            sel_bg = "#EFF6FF"          # Soft blue highlight tint
            sel_fg = "#1D4ED8"          # Deep blue highlighted text
            scroll_bar = "#EBF0F7"
            scroll_handle = "#CBD5E1"
            scroll_hover = "#94A3B8"

            cb_off = assets.get("cb_off_light", assets["cb_off"])
            cb_hover = assets.get("cb_hover_light", assets["cb_hover"])
            cb_on = assets.get("cb_on_light", assets["cb_on"])
            rb_off = assets.get("rb_off_light", assets["rb_off"])
            rb_hover = assets.get("rb_hover_light", assets["rb_hover"])
            rb_on = assets.get("rb_on_light", assets["rb_on"])
        else:
            # Professional Dark Forensic Workstation Theme (Reference Image 2: Materio Dark)
            bg_main = "#0b0e17"
            bg_sidebar = "#111625"
            bg_header = "#111625"
            bg_card = "#161b2b"
            bg_input = "#0e121d"
            bg_table = "#0e121d"
            bg_head = "#161b2b"
            border_color = "#232b43"
            border_strong = "#2e3654"
            border_subtle = "#1c2236"
            text_primary = "#f1f5f9"
            text_secondary = "#8e98b0"
            accent_blue = "#3b82f6"
            accent_hover = "#60a5fa"
            sel_bg = "#1e3a8a"
            sel_fg = "#ffffff"
            scroll_bar = "#0e121d"
            scroll_handle = "#232b43"
            scroll_hover = "#3b82f6"

            cb_off = assets["cb_off"]
            cb_hover = assets["cb_hover"]
            cb_on = assets["cb_on"]
            rb_off = assets["rb_off"]
            rb_hover = assets["rb_hover"]
            rb_on = assets["rb_on"]

        # Synchronize Application Palette
        pal = self.palette()
        pal.setColor(QPalette.Window, QColor(bg_main))
        pal.setColor(QPalette.Base, QColor(bg_table))
        pal.setColor(QPalette.AlternateBase, QColor(bg_table))
        pal.setColor(QPalette.Text, QColor(text_primary))
        pal.setColor(QPalette.WindowText, QColor(text_primary))
        self.setPalette(pal)

        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {bg_main};
            }}
            QWidget {{
                color: {text_primary};
                font-family: 'Segoe UI', -apple-system, sans-serif;
            }}
            #SidebarFrame {{
                background-color: {bg_sidebar};
                border-right: 1px solid {border_color};
            }}
            #TopHeader {{
                background-color: {bg_header};
                border-bottom: 1px solid {border_color};
            }}
            #MainContentContainer, QStackedWidget, QStackedWidget > QWidget, QScrollArea, QScrollArea > QWidget {{
                background-color: {bg_main};
            }}
            #SidebarBrandLogoTile {{
                background-color: #111625;
                border: none;
                border-radius: 6px;
            }}
            #SidebarBrandTitle {{
                font-size: 13px;
                font-weight: 800;
                letter-spacing: 0.5px;
                color: {text_primary};
            }}
            #SidebarBrandSubtitle {{
                font-size: 10px;
                font-weight: 500;
                color: {text_secondary};
            }}
            #SidebarSectionLabel {{
                font-size: 10px;
                font-weight: 700;
                color: {text_secondary};
                letter-spacing: 1px;
                padding: 6px 4px 2px 4px;
            }}
            #SidebarVersionLabel {{
                color: {text_secondary};
                font-size: 10px;
                font-weight: 500;
            }}
            #SidebarDivider {{
                background-color: {border_color};
            }}
            QScrollArea, QScrollArea > QWidget, QScrollArea > QWidget > QWidget, QStackedWidget, #MainContentContainer {{
                background-color: {bg_main};
                border: none;
            }}
            QScrollBar:vertical {{
                background: {scroll_bar};
                width: 10px;
                border-radius: 5px;
                margin: 0px;
            }}
            QScrollBar::handle:vertical {{
                background: {scroll_handle};
                min-height: 25px;
                border-radius: 5px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {scroll_hover};
            }}
            QScrollBar:horizontal {{
                background: {scroll_bar};
                height: 10px;
                border-radius: 5px;
                margin: 0px;
            }}
            QScrollBar::handle:horizontal {{
                background: {scroll_handle};
                min-width: 25px;
                border-radius: 5px;
            }}
            QScrollBar::handle:horizontal:hover {{
                background: {scroll_hover};
            }}
            QScrollBar::add-line, QScrollBar::sub-line {{
                width: 0px;
                height: 0px;
            }}
            QLineEdit {{
                background-color: {bg_input};
                border: 1px solid {border_color};
                border-radius: 8px;
                color: {text_primary};
                padding: 8px 12px;
                font-size: 12px;
            }}
            QLineEdit:focus {{
                border-color: {accent_blue};
            }}
            QCheckBox {{
                spacing: 10px;
                font-size: 12px;
                color: {text_primary};
            }}
            QCheckBox::indicator {{
                width: 20px;
                height: 20px;
            }}
            QCheckBox::indicator:unchecked {{
                image: url({cb_off});
            }}
            QCheckBox::indicator:unchecked:hover {{
                image: url({cb_hover});
            }}
            QCheckBox::indicator:checked {{
                image: url({cb_on});
            }}
            QRadioButton {{
                spacing: 10px;
                font-size: 12px;
                color: {text_primary};
            }}
            QRadioButton::indicator {{
                width: 20px;
                height: 20px;
            }}
            QRadioButton::indicator:unchecked {{
                image: url({rb_off});
            }}
            QRadioButton::indicator:unchecked:hover {{
                image: url({rb_hover});
            }}
            QRadioButton::indicator:checked {{
                image: url({rb_on});
            }}
            QComboBox {{
                background-color: {bg_input};
                border: 1px solid {border_color};
                border-radius: 8px;
                color: {text_primary};
                padding: 6px 12px;
                font-size: 12px;
                min-width: 140px;
            }}
            QComboBox:focus {{
                border-color: {accent_blue};
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {bg_card};
                border: 1px solid {border_color};
                color: {text_primary};
                selection-background-color: {accent_blue};
                selection-color: #ffffff;
                outline: none;
            }}
            QProgressBar {{
                background-color: {bg_input};
                border: 1px solid {border_color};
                border-radius: 6px;
                text-align: center;
                color: {text_primary};
                font-size: 11px;
                font-weight: 600;
                height: 16px;
            }}
            QProgressBar::chunk {{
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {accent_blue}, stop:1 {accent_hover});
                border-radius: 5px;
            }}
            QTableWidget {{
                background-color: {bg_table};
                border: 1px solid {border_color};
                border-radius: 8px;
                gridline-color: {border_subtle};
                color: {text_primary};
                font-size: 12px;
                selection-background-color: {sel_bg};
                selection-color: {sel_fg};
            }}
            QTableWidget::item {{
                padding: 6px 8px;
            }}
            QTableWidget::item:selected {{
                background-color: {sel_bg};
                color: {sel_fg};
            }}
            QTreeWidget {{
                background-color: {bg_table};
                alternate-background-color: {bg_table};
                border: 1px solid {border_color};
                border-radius: 8px;
                color: {text_primary};
                font-size: 12px;
                selection-background-color: {sel_bg};
                selection-color: {sel_fg};
            }}
            QTreeWidget::item {{
                background-color: {bg_table};
                color: {text_primary};
                padding: 6px 8px;
                border-bottom: 1px solid {border_subtle};
            }}
            QTreeWidget::item:hover {{
                background-color: {sel_bg};
                color: {sel_fg};
            }}
            QTreeWidget::item:selected {{
                background-color: {sel_bg};
                color: {sel_fg};
            }}
            QHeaderView::section {{
                background-color: {bg_head};
                color: {text_secondary};
                font-size: 11px;
                font-weight: 700;
                text-transform: uppercase;
                border: none;
                border-bottom: 1px solid {border_color};
                padding: 8px;
            }}
        """)

        # Style Header Titles
        if hasattr(self, "header_title") and self.header_title:
            self.header_title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {text_primary}; background: transparent; border: none;")
        if hasattr(self, "header_subtitle") and self.header_subtitle:
            self.header_subtitle.setStyleSheet(f"font-size: 11px; color: {text_secondary}; background: transparent; border: none;")

        # Style Air-Gapped Header Badge
        if hasattr(self, "header_air_gapped") and self.header_air_gapped:
            if theme == "light":
                self.header_air_gapped.setStyleSheet("""
                    background-color: rgba(37, 99, 235, 0.08);
                    color: #1d4ed8;
                    border: 1px solid rgba(37, 99, 235, 0.25);
                    border-radius: 6px;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: 700;
                    letter-spacing: 0.5px;
                """)
            else:
                self.header_air_gapped.setStyleSheet("""
                    background-color: rgba(59, 130, 246, 0.12);
                    color: #60a5fa;
                    border: 1px solid rgba(59, 130, 246, 0.35);
                    border-radius: 6px;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: 700;
                    letter-spacing: 0.5px;
                """)

        # Style Header Compare Button
        if hasattr(self, "header_compare_btn") and self.header_compare_btn:
            cmp_ic_col = "#2563eb" if theme == "light" else "#60a5fa"
            self.header_compare_btn.setIcon(LucideIcons.icon("compare", color=cmp_ic_col, size=16))
            self.header_compare_btn.setIconSize(QSize(16, 16))
            if theme == "light":
                self.header_compare_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #ffffff;
                        border: 1px solid #cbd5e1;
                        color: #0f172a;
                        border-radius: 8px;
                        padding: 7px 14px;
                        font-size: 12px;
                        font-weight: 600;
                    }
                    QPushButton:hover {
                        background-color: #f1f5f9;
                        border-color: #2563eb;
                    }
                """)
            else:
                self.header_compare_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #1e243b;
                        border: 1px solid #2e3654;
                        color: #f1f5f9;
                        border-radius: 8px;
                        padding: 7px 14px;
                        font-size: 12px;
                        font-weight: 600;
                    }
                    QPushButton:hover {
                        background-color: #28304e;
                        border-color: #3b82f6;
                    }
                """)

        # Style Header Theme Button
        if hasattr(self, "header_theme_btn") and self.header_theme_btn:
            if theme == "light":
                self.header_theme_btn.setIcon(LucideIcons.icon("moon", color="#3b82f6", size=18))
                self.header_theme_btn.setToolTip("Switch to Dark Forensic Mode")
                self.header_theme_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #f1f5f9;
                        border: 1px solid #cbd5e1;
                        border-radius: 8px;
                    }
                    QPushButton:hover {
                        background-color: #e2e8f0;
                        border-color: #2563eb;
                    }
                """)
            else:
                self.header_theme_btn.setIcon(LucideIcons.icon("sun", color="#f59e0b", size=18))
                self.header_theme_btn.setToolTip("Switch to Clean Light Mode")
                self.header_theme_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #1e243b;
                        border: 1px solid #2e3654;
                        border-radius: 8px;
                    }
                    QPushButton:hover {
                        background-color: #28304e;
                        border-color: #3b82f6;
                    }
                """)

        # Style Sidebar Theme Button
        if hasattr(self, "sidebar_theme_btn") and self.sidebar_theme_btn:
            if theme == "light":
                self.sidebar_theme_btn.setIcon(LucideIcons.icon("moon", color="#3b82f6", size=16))
                self.sidebar_theme_btn.setText("  Dark Mode")
                self.sidebar_theme_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #f8fafc;
                        border: 1px solid #e2e8f0;
                        border-radius: 8px;
                        color: #334155;
                        font-size: 12px;
                        font-weight: 600;
                        text-align: left;
                        padding: 6px 12px;
                    }
                    QPushButton:hover {
                        background-color: #f1f5f9;
                        border-color: #2563eb;
                        color: #0f172a;
                    }
                """)
            else:
                self.sidebar_theme_btn.setIcon(LucideIcons.icon("sun", color="#f59e0b", size=16))
                self.sidebar_theme_btn.setText("  Light Mode")
                self.sidebar_theme_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #161c2e;
                        border: 1px solid #232b43;
                        border-radius: 8px;
                        color: #e2e8f0;
                        font-size: 12px;
                        font-weight: 600;
                        text-align: left;
                        padding: 6px 12px;
                    }
                    QPushButton:hover {
                        background-color: #1e243b;
                        border-color: #3b82f6;
                        color: #ffffff;
                    }
                """)

        # Style Sidebar Engine Status Pill
        if hasattr(self, "sidebar_engine_pill") and self.sidebar_engine_pill:
            if theme == "light":
                pill_col = "#334155" if self.exiftool_ok else "#dc2626"
                self.sidebar_engine_pill.setStyleSheet(f"""
                    background-color: #f1f5f9;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    padding: 5px 10px;
                    font-size: 11px;
                    font-weight: 600;
                    color: {pill_col};
                """)
            else:
                pill_col = "#94a3b8" if self.exiftool_ok else "#f87171"
                self.sidebar_engine_pill.setStyleSheet(f"""
                    background-color: #161c2e;
                    border: 1px solid #232b43;
                    border-radius: 6px;
                    padding: 5px 10px;
                    font-size: 11px;
                    font-weight: 600;
                    color: {pill_col};
                """)

        # Update Overview Hero Card & Typography
        if hasattr(self, "overview_modules_hdr") and self.overview_modules_hdr:
            self.overview_modules_hdr.setStyleSheet(f"font-size: 11px; font-weight: 700; letter-spacing: 1.2px; color: {text_secondary}; text-transform: uppercase; margin-top: 2px; margin-bottom: 16px;")

        if hasattr(self, "overview_empty_log_lbl") and self.overview_empty_log_lbl:
            if theme == "light":
                self.overview_empty_log_lbl.setStyleSheet("color: #64748b; font-size: 12px; padding: 24px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px;")
            else:
                self.overview_empty_log_lbl.setStyleSheet("color: #8e98b0; font-size: 12px; padding: 24px; background: rgba(14, 18, 29, 0.4); border-radius: 8px;")

        # Update Sidebar Nav Buttons
        if hasattr(self, "sidebar_nav_buttons"):
            for i, btn in enumerate(self.sidebar_nav_buttons):
                btn.set_active(i == getattr(self, "active_nav_index", 0), theme=theme)

        # Recursively apply theme updates to all SurfaceCards, StatCards, FileDropWidgets, and MediaPreviews
        for card in self.findChildren(SurfaceCard):
            card.apply_theme(theme)
        if hasattr(self, "meta_preview_gallery") and self.meta_preview_gallery:
            self.meta_preview_gallery.apply_theme(theme)
        for stat in self.findChildren(ModernStatCard):
            stat.apply_theme(theme)
        for drop in self.findChildren(FileDropWidget):
            drop.apply_theme(theme)
        for prev in self.findChildren(MediaPreviewWidget):
            prev.apply_theme(theme)
        for gal in self.findChildren(SmallGalleryView):
            gal.apply_theme(theme)

        # Update Workstation Dynamic Labels & Indicators for High Contrast
        if hasattr(self, "inv_status_lbl") and self.inv_status_lbl:
            self.inv_status_lbl.setStyleSheet(f"font-size: 13px; font-weight: 700; color: {text_primary};")
        if hasattr(self, "inv_ticker_lbl") and self.inv_ticker_lbl:
            self.inv_ticker_lbl.setStyleSheet(f"font-family: monospace; font-size: 11px; color: {text_secondary};")
        if hasattr(self, "results_desc") and self.results_desc:
            self.results_desc.setStyleSheet(f"font-size: 11px; color: {text_secondary};")
        if hasattr(self, "results_btn_layout"):
            self._style_results_actions(theme)
        if hasattr(self, "hash_details_hint"):
            self.hash_details_hint.setStyleSheet(f"font-size: 11px; color: {text_secondary}; background: transparent;")
            self.hash_details_hint_icon.setPixmap(LucideIcons.pixmap("info", color=text_secondary, size=16))
        if hasattr(self, "inv_fmt_lbl") and self.inv_fmt_lbl:
            self.inv_fmt_lbl.setStyleSheet(f"font-size: 11px; font-weight: 700; color: {text_secondary};")
        if hasattr(self, "inv_out_path_lbl") and self.inv_out_path_lbl:
            self.inv_out_path_lbl.setStyleSheet(f"font-size: 11px; font-weight: 700; color: {text_secondary};")
        if hasattr(self, "hash_target_lbl") and self.hash_target_lbl:
            self.hash_target_lbl.setStyleSheet(f"font-size: 11px; font-weight: 700; color: {text_secondary};")
        if hasattr(self, "hash_export_lbl") and self.hash_export_lbl:
            self.hash_export_lbl.setStyleSheet(f"font-size: 11px; font-weight: 700; color: {text_secondary}; margin-top: 10px;")

        # Update Metadata Viewer Badges and Status Pills
        if hasattr(self, "meta_status_badge") and self.meta_status_badge:
            if theme == "light":
                self.meta_status_badge.setStyleSheet("font-size: 11px; font-weight: 700; color: #1d4ed8; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 4px 10px;")
            else:
                self.meta_status_badge.setStyleSheet("font-size: 11px; font-weight: 700; color: #93c5fd; background: rgba(59, 130, 246, 0.12); border: 1px solid rgba(59, 130, 246, 0.25); border-radius: 6px; padding: 4px 10px;")

        if hasattr(self, "meta_count_pill") and self.meta_count_pill:
            if theme == "light":
                self.meta_count_pill.setStyleSheet("font-size: 11px; font-weight: 600; color: #0f172a; background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 6px; padding: 4px 8px;")
            else:
                self.meta_count_pill.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0; background: rgba(14, 18, 29, 0.8); border: 1px solid #232b43; border-radius: 6px; padding: 4px 8px;")

        if hasattr(self, "meta_integrity_pill") and self.meta_integrity_pill:
            if theme == "light":
                self.meta_integrity_pill.setStyleSheet("font-size: 11px; font-weight: 600; color: #0f172a; background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 6px; padding: 4px 8px;")
            else:
                self.meta_integrity_pill.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0; background: rgba(14, 18, 29, 0.8); border: 1px solid #232b43; border-radius: 6px; padding: 4px 8px;")

        if hasattr(self, "meta_tags_count_lbl") and self.meta_tags_count_lbl:
            if theme == "light":
                self.meta_tags_count_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #1d4ed8; padding-left: 6px;")
            else:
                self.meta_tags_count_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #93c5fd; padding-left: 6px;")

        # Update Processed Output Folder Button
        if hasattr(self, "meta_btn_open_processed_folder") and self.meta_btn_open_processed_folder:
            if theme == "light":
                self.meta_btn_open_processed_folder.setStyleSheet("""
                    QPushButton {
                        background-color: #ffffff;
                        border: 1.5px solid #2563eb;
                        border-radius: 8px;
                        color: #1d4ed8;
                        padding: 8px 16px;
                        font-size: 12px;
                        font-weight: 700;
                    }
                    QPushButton:hover {
                        background-color: #eff6ff;
                        border-color: #1d4ed8;
                    }
                """)
            else:
                self.meta_btn_open_processed_folder.setStyleSheet("""
                    QPushButton {
                        background-color: #1e243b;
                        border: 1.5px solid #3b82f6;
                        border-radius: 8px;
                        color: #ffffff;
                        padding: 8px 16px;
                        font-size: 12px;
                        font-weight: 700;
                    }
                    QPushButton:hover {
                        background-color: #2563eb;
                        border-color: #60a5fa;
                    }
                """)

        # Update Metadata Viewer Export and Raw JSON buttons
        viewer_btns = [
            getattr(self, "meta_btn_export_viewer", None),
            getattr(self, "meta_btn_export_html_viewer", None),
            getattr(self, "meta_btn_view_raw_json", None),
        ]
        for btn in set(b for b in viewer_btns if b is not None):
            self._style_button_secondary(btn, theme)
        if hasattr(self, "meta_btn_export_viewer") and self.meta_btn_export_viewer:
            ic_c = "#1d4ed8" if theme == "light" else "#93c5fd"
            self.meta_btn_export_viewer.setIcon(LucideIcons.icon("download", color=ic_c, size=14))

        # Update All Workstation Buttons with High Contrast Palette
        self._update_all_workstation_buttons(theme)

        # Update Overview Action Buttons
        if hasattr(self, "overview_action_buttons"):
            for btn in self.overview_action_buttons:
                btn.setStyleSheet(self._action_btn_qss())

        # Update Overview Descriptions & Bullets
        desc_color = "#475569" if theme == "light" else "#8e98b0"
        if hasattr(self, "overview_desc_labels"):
            for lbl in self.overview_desc_labels:
                lbl.setStyleSheet(f"font-size: 11px; line-height: 1.35; color: {desc_color};")

        dot_color = "#2563eb" if theme == "light" else "#60a5fa"
        txt_color = "#334155" if theme == "light" else "#cbd5e1"
        if hasattr(self, "overview_bullet_labels"):
            for dot, blbl in self.overview_bullet_labels:
                dot.setStyleSheet(f"color: {dot_color}; font-size: 13px; font-weight: 900;")
                blbl.setStyleSheet(f"font-size: 11px; color: {txt_color}; font-weight: 500;")

        if hasattr(self, "overview_modules_hdr") and self.overview_modules_hdr:
            self.overview_modules_hdr.setStyleSheet("font-size: 11px; font-weight: 700; letter-spacing: 1.2px; color: #64748b; text-transform: uppercase; margin-top: 2px; margin-bottom: 16px;" if theme == "light" else "font-size: 11px; font-weight: 700; letter-spacing: 1.2px; color: #8e98b0; text-transform: uppercase; margin-top: 2px; margin-bottom: 16px;")

        if hasattr(self, "overview_empty_log_lbl") and self.overview_empty_log_lbl:
            empty_bg = "#f8fafc" if theme == "light" else "rgba(14, 18, 29, 0.4)"
            empty_color = "#64748b" if theme == "light" else "#8e98b0"
            empty_border = "#cbd5e1" if theme == "light" else "#232b43"
            self.overview_empty_log_lbl.setStyleSheet(f"color: {empty_color}; font-size: 11px; padding: 24px; background: {empty_bg}; border: 1px dashed {empty_border}; border-radius: 8px;")

        # Update Settings Buttons
        if hasattr(self, "settings_buttons"):
            for btn in self.settings_buttons:
                btn.setStyleSheet(self._settings_btn_qss())

        # Update Metadata Tree Buttons
        if hasattr(self, "meta_btn_expand") and self.meta_btn_expand:
            if theme == "light":
                tree_btn_css = """
                    QPushButton {
                        background-color: #f1f5f9;
                        border: 1px solid #cbd5e1;
                        border-radius: 6px;
                        color: #0f172a;
                        padding: 6px 14px;
                        font-size: 11px;
                        font-weight: 600;
                    }
                    QPushButton:hover {
                        background-color: #e2e8f0;
                        border-color: #2563eb;
                        color: #1e40af;
                    }
                """
            else:
                tree_btn_css = """
                    QPushButton {
                        background-color: #1e243b;
                        border: 1px solid #2e3654;
                        border-radius: 6px;
                        color: #f1f5f9;
                        padding: 6px 14px;
                        font-size: 11px;
                        font-weight: 600;
                    }
                    QPushButton:hover {
                        background-color: #28304e;
                        border-color: #3b82f6;
                    }
                """
            self.meta_btn_expand.setStyleSheet(tree_btn_css)
            if hasattr(self, "meta_btn_export_viewer") and self.meta_btn_export_viewer:
                self.meta_btn_export_viewer.setStyleSheet(tree_btn_css)
            if hasattr(self, "meta_btn_collapse") and self.meta_btn_collapse:
                self.meta_btn_collapse.setStyleSheet(tree_btn_css)
            if hasattr(self, "meta_btn_copy") and self.meta_btn_copy:
                self.meta_btn_copy.setStyleSheet(tree_btn_css)
            if hasattr(self, "meta_btn_view_json") and self.meta_btn_view_json:
                self.meta_btn_view_json.setStyleSheet(tree_btn_css)

        # Update Metadata Browse Buttons
        if hasattr(self, "meta_btn_browse") and self.meta_btn_browse:
            if theme == "light":
                meta_brw_css = """
                    QPushButton {
                        background-color: #f1f5f9;
                        border: 1px solid #cbd5e1;
                        border-radius: 8px;
                        color: #0f172a;
                        padding: 8px 16px;
                        font-weight: 600;
                    }
                    QPushButton:hover { background-color: #e2e8f0; border-color: #2563eb; }
                """
            else:
                meta_brw_css = """
                    QPushButton {
                        background-color: #1e243b;
                        border: 1px solid #2e3654;
                        border-radius: 8px;
                        color: #f1f5f9;
                        padding: 8px 16px;
                        font-weight: 600;
                    }
                    QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
                """
            self.meta_btn_browse.setStyleSheet(meta_brw_css)
            if hasattr(self, "meta_btn_browse_out") and self.meta_btn_browse_out:
                self.meta_btn_browse_out.setStyleSheet(meta_brw_css)

        # Update metadata left variables frame & labels
        if hasattr(self, "meta_vars_frame") and self.meta_vars_frame:
            if theme == "light":
                self.meta_vars_frame.setStyleSheet("""
                    #MetaVarsFrame {
                        background-color: #f8fafc;
                        border: 1px solid #cbd5e1;
                        border-radius: 8px;
                    }
                """)
                if hasattr(self, "meta_info_lbl") and self.meta_info_lbl:
                    self.meta_info_lbl.setStyleSheet("""
                        font-size: 11px;
                        color: #334155;
                        background-color: #f1f5f9;
                        border: 1px solid #e2e8f0;
                        border-radius: 6px;
                        padding: 6px 10px;
                    """)
                if hasattr(self, "meta_left_progress_bar") and self.meta_left_progress_bar:
                    self.meta_left_progress_bar.setStyleSheet("""
                        QProgressBar { background-color: #e2e8f0; border: none; border-radius: 2px; }
                        QProgressBar::chunk { background-color: #2563eb; border-radius: 2px; }
                    """)
                if hasattr(self, "meta_proc_status_lbl") and self.meta_proc_status_lbl:
                    self.meta_proc_status_lbl.setStyleSheet("font-size: 11px; color: #475569;")
                if hasattr(self, "meta_left_count_lbl") and self.meta_left_count_lbl:
                    self.meta_left_count_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #475569;")
            else:
                self.meta_vars_frame.setStyleSheet("""
                    #MetaVarsFrame {
                        background-color: rgba(14, 18, 29, 0.6);
                        border: 1px solid #232b43;
                        border-radius: 8px;
                    }
                """)
                if hasattr(self, "meta_info_lbl") and self.meta_info_lbl:
                    self.meta_info_lbl.setStyleSheet("""
                        font-size: 11px;
                        color: #8e98b0;
                        background-color: rgba(14, 18, 29, 0.6);
                        border: 1px solid #232b43;
                        border-radius: 6px;
                        padding: 6px 10px;
                    """)
                if hasattr(self, "meta_left_progress_bar") and self.meta_left_progress_bar:
                    self.meta_left_progress_bar.setStyleSheet("""
                        QProgressBar { background-color: #111422; border: none; border-radius: 2px; }
                        QProgressBar::chunk { background-color: #3b82f6; border-radius: 2px; }
                    """)
                if hasattr(self, "meta_proc_status_lbl") and self.meta_proc_status_lbl:
                    self.meta_proc_status_lbl.setStyleSheet("font-size: 11px; color: #64748b;")
                if hasattr(self, "meta_left_count_lbl") and self.meta_left_count_lbl:
                    self.meta_left_count_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0;")

        if hasattr(self, "_update_submode_buttons"):
            self._update_submode_buttons()

    def _build_ui(self):
        """Construct modern 2-column forensic workstation shell: Sidebar + Top Header + Stacked Pages."""
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # -------------------------------------------------------------
        # LEFT COLUMN: COMPACT SIDEBAR NAVIGATION
        # -------------------------------------------------------------
        self.sidebar_frame = QFrame()
        self.sidebar_frame.setObjectName("SidebarFrame")
        self.sidebar_frame.setFixedWidth(230)
        sidebar_layout = QVBoxLayout(self.sidebar_frame)
        sidebar_layout.setContentsMargins(14, 16, 14, 16)
        sidebar_layout.setSpacing(8)

        # Branding Header (App Logo + Title)
        brand_container = QHBoxLayout()
        brand_container.setSpacing(10)
        brand_container.setContentsMargins(4, 2, 4, 8)

        brand_logo = QLabel()
        brand_logo.setObjectName('SidebarBrandLogoTile')
        brand_logo.setFixedSize(30, 30)
        brand_logo.setAlignment(Qt.AlignCenter)
        self.sidebar_brand_logo = brand_logo
        if self.assets.get("logo") and os.path.exists(self.assets["logo"]):
            pm = QPixmap(self.assets["logo"])
            if not pm.isNull():
                brand_logo.setPixmap(pm.scaled(30, 30, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            brand_logo.setPixmap(LucideIcons.pixmap("shield", color="#3b82f6", size=26))
        brand_container.addWidget(brand_logo)

        brand_titles = QVBoxLayout()
        brand_titles.setSpacing(0)
        self.sidebar_brand_title = QLabel(APP_NAME)
        self.sidebar_brand_title.setObjectName("SidebarBrandTitle")
        self.sidebar_brand_subtitle = QLabel(APP_SUBTITLE)
        self.sidebar_brand_subtitle.setObjectName("SidebarBrandSubtitle")
        brand_titles.addWidget(self.sidebar_brand_title)
        brand_titles.addWidget(self.sidebar_brand_subtitle)
        brand_container.addLayout(brand_titles)
        brand_container.addStretch()

        sidebar_layout.addLayout(brand_container)

        # Divider
        sidebar_divider1 = QFrame()
        sidebar_divider1.setFrameShape(QFrame.HLine)
        sidebar_divider1.setObjectName("SidebarDivider")
        sidebar_divider1.setFixedHeight(1)
        sidebar_layout.addWidget(sidebar_divider1)
        sidebar_layout.addSpacing(6)

        # Section Label: MODULES
        sec_modules = QLabel("WORKSTATION MODULES")
        sec_modules.setObjectName("SidebarSectionLabel")
        sidebar_layout.addWidget(sec_modules)

        # Module Navigation Buttons
        self.btn_nav_overview = SidebarNavButton("Overview", "home")
        self.btn_nav_hash = SidebarNavButton("Hash Verification", "shield")
        self.btn_nav_inventory = SidebarNavButton("File Inventory", "folder")
        self.btn_nav_metadata = SidebarNavButton("Metadata Explorer", "search")

        self.btn_nav_overview.clicked.connect(self._switch_to_overview_mode)
        self.btn_nav_hash.clicked.connect(self._switch_to_hash_mode)
        self.btn_nav_inventory.clicked.connect(self._switch_to_inventory_mode)
        self.btn_nav_metadata.clicked.connect(self._switch_to_metadata_mode)

        sidebar_layout.addWidget(self.btn_nav_overview)
        sidebar_layout.addWidget(self.btn_nav_hash)
        sidebar_layout.addWidget(self.btn_nav_inventory)
        sidebar_layout.addWidget(self.btn_nav_metadata)

        # Spacer
        sidebar_layout.addStretch()

        # Section Label: PREFERENCES / SYSTEM
        sec_pref = QLabel("SYSTEM & PREFERENCES")
        sec_pref.setObjectName("SidebarSectionLabel")
        sidebar_layout.addWidget(sec_pref)

        # Settings Button
        self.btn_nav_settings = SidebarNavButton("Settings", "settings")
        self.btn_nav_settings.clicked.connect(self._switch_to_settings_mode)
        sidebar_layout.addWidget(self.btn_nav_settings)

        # Sidebar Theme Toggle Button
        self.sidebar_theme_btn = QPushButton("  Light Mode")
        self.sidebar_theme_btn.setObjectName("SidebarThemeBtn")
        self.sidebar_theme_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.sidebar_theme_btn.setFixedHeight(36)
        self.sidebar_theme_btn.clicked.connect(self._toggle_theme)
        sidebar_layout.addWidget(self.sidebar_theme_btn)

        sidebar_layout.addSpacing(4)
        sidebar_divider2 = QFrame()
        sidebar_divider2.setFrameShape(QFrame.HLine)
        sidebar_divider2.setObjectName("SidebarDivider")
        sidebar_divider2.setFixedHeight(1)
        sidebar_layout.addWidget(sidebar_divider2)
        sidebar_layout.addSpacing(4)

        # Engine Status Pill in Sidebar Bottom
        self.sidebar_engine_pill = QLabel()
        self.sidebar_engine_pill.setObjectName("SidebarEnginePill")
        self.sidebar_engine_pill.setAlignment(Qt.AlignCenter)
        if self.exiftool_ok:
            self.sidebar_engine_pill.setText(f"ExifTool v{self.exiftool_ver}")
        else:
            self.sidebar_engine_pill.setText("ExifTool Not Found")
        sidebar_layout.addWidget(self.sidebar_engine_pill)

        # Workstation Version Footer
        self.sidebar_version_lbl = QLabel(f"v{APP_VERSION} • {APP_TAGLINE}")
        self.sidebar_version_lbl.setObjectName("SidebarVersionLabel")
        self.sidebar_version_lbl.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(self.sidebar_version_lbl)

        root_layout.addWidget(self.sidebar_frame)

        # -------------------------------------------------------------
        # RIGHT COLUMN: MAIN CONTENT CONTAINER
        # -------------------------------------------------------------
        self.main_content_container = QWidget()
        self.main_content_container.setObjectName("MainContentContainer")
        content_vlayout = QVBoxLayout(self.main_content_container)
        content_vlayout.setContentsMargins(10, 10, 10, 10)
        content_vlayout.setSpacing(0)
        # Top Header Bar removed as requested - all actions accessible via sidebar & workspace
        self.top_header = None
        self.header_title = None
        self.header_subtitle = None
        self.header_air_gapped = None
        self.header_compare_btn = None
        self.header_theme_btn = None

        # Main Workspace: Stacked Widget with 5 Pages
        self.stacked_pages = QStackedWidget()
        self.stacked_pages.setContentsMargins(8, 6, 8, 8)

        self.page_overview = self._create_overview_page()
        self.page_hash = self._create_hash_page()
        self.page_inventory = self._create_inventory_page()
        self.page_metadata = self._create_metadata_page()
        self.page_settings = self._create_settings_page()

        self.stacked_pages.addWidget(self.page_overview)    # 0
        self.stacked_pages.addWidget(self.page_hash)        # 1
        self.stacked_pages.addWidget(self.page_inventory)   # 2
        self.stacked_pages.addWidget(self.page_metadata)    # 3
        self.stacked_pages.addWidget(self.page_settings)    # 4

        content_vlayout.addWidget(self.stacked_pages)
        root_layout.addWidget(self.main_content_container)

        self.sidebar_nav_buttons = [
            self.btn_nav_overview,
            self.btn_nav_hash,
            self.btn_nav_inventory,
            self.btn_nav_metadata,
            self.btn_nav_settings,
        ]

    # ==========================================================================
    # NAVIGATION & THEME SWITCHING
    # ==========================================================================
    def _set_active_nav_index(self, index: int):
        """Switch active page and update sidebar states, header titles, and QSettings."""
        if not hasattr(self, "stacked_pages") or not (0 <= index < self.stacked_pages.count()):
            return
        self.active_nav_index = index
        self.stacked_pages.setCurrentIndex(index)

        # Update sidebar button states
        if hasattr(self, "sidebar_nav_buttons"):
            for i, btn in enumerate(self.sidebar_nav_buttons):
                btn.set_active(i == index, theme=self.current_theme)

        # Header Titles & Subtitles mapping
        headers = [
            ("Workstation Overview", "Central Launchpad • Module Access & Operational Telemetry"),
            ("Hash Verification", "Cryptographic Integrity • Multi-Algorithm Verification & Tamper Detection"),
            ("File Inventory & Directory Audit", "Automated Recursive Cataloging • Media Metadata Extraction & Reports"),
            ("Forensic Metadata Explorer", "Official ExifTool Integration • EXIF / XMP / IPTC Extraction & Pre/Post Integrity"),
            ("Workstation Settings", "Engine Paths • Default Hash Suites • Report Directories & Interface Preferences")
        ]
        if getattr(self, "header_title", None) and getattr(self, "header_subtitle", None) and index < len(headers):
            title, subtitle = headers[index]
            self.header_title.setText(title)
            self.header_subtitle.setText(subtitle)

        # Synchronize legacy active_mode string for backwards compatibility
        modes = ["overview", "hash", "inventory", "metadata", "settings"]
        if index < len(modes):
            self.active_mode = modes[index]

    def _switch_to_overview_mode(self):
        self._set_active_nav_index(0)

    def _switch_to_hash_mode(self):
        self._set_active_nav_index(1)

    def _switch_to_inventory_mode(self):
        self._set_active_nav_index(2)

    def _switch_to_metadata_mode(self):
        self._set_active_nav_index(3)

    def _switch_to_settings_mode(self):
        self._set_active_nav_index(4)

    def _update_mode_button_styles(self):
        """Refresh sidebar button appearance according to active index and current theme."""
        if hasattr(self, "sidebar_nav_buttons"):
            for i, btn in enumerate(self.sidebar_nav_buttons):
                btn.set_active(i == getattr(self, "active_nav_index", 0), theme=self.current_theme)

    def _toggle_theme(self):
        """Toggle between Dark and Light mode across all application pages."""
        new_theme = "light" if self.current_theme == "dark" else "dark"
        if hasattr(self, "settings_rb_theme_light") and hasattr(self, "settings_rb_theme_dark"):
            if new_theme == "light":
                self.settings_rb_theme_light.setChecked(True)
            else:
                self.settings_rb_theme_dark.setChecked(True)
        self._apply_global_theme(new_theme)

    def _open_compare_dialog(self):
        dlg = CompareTwoFilesDialog(parent=self)
        dlg.exec()

    # ==========================================================================
    # OVERVIEW LAUNCHPAD PAGE
    # ==========================================================================
        # ==========================================================================
    # OVERVIEW LAUNCHPAD PAGE
    # ==========================================================================
    def _create_overview_page(self):
        """
        Constructs the spacious Overview Launchpad screen:
        1. Top Row: 3 Spacious Core Module Shortcut Cards (bounded height 240-290px, eliminating excessive vertical void on widescreen).
        2. Middle Row: Quick Actions (spacious vertical button stack) + Workstation Environment (key-value telemetry).
        3. Bottom Row: Session Activity Log (expands with stretch 1 to show real operational records).
        """
        theme = getattr(self, "current_theme", "dark")
        self.overview_shortcut_buttons = []
        self.overview_desc_labels = []
        self.overview_bullet_labels = []

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(6, 6, 6, 24)
        layout.setSpacing(18)

        # Section Header
        self.overview_modules_hdr = QLabel("CORE WORKSTATION MODULES")
        self.overview_modules_hdr.setStyleSheet("font-size: 11px; font-weight: 700; letter-spacing: 1.2px; color: #64748b; text-transform: uppercase; margin-top: 2px; margin-bottom: 16px;" if theme == "light" else "font-size: 11px; font-weight: 700; letter-spacing: 1.2px; color: #8e98b0; text-transform: uppercase; margin-top: 2px; margin-bottom: 16px;")
        layout.addWidget(self.overview_modules_hdr)

        # 1. THREE CORE MODULE CARDS
        tiles_row = QHBoxLayout()
        tiles_row.setSpacing(16)

        tile_hash = self._create_module_shortcut_card(
            title="Hash Verification",
            icon="shield",
            description="Calculate SHA-256, SHA-512, MD5, and SHA-1 checksums in a single high-speed read pass. Verify against expected benchmark manifests or compare two files side-by-side with immediate tamper alerts.",
            bullets=["Multi-hash parallel calculation", "Side-by-side file comparison", "Checksum export (sha256/md5)"],
            btn_text="Open Hash Verification",
            on_click=self._switch_to_hash_mode
        )
        tiles_row.addWidget(tile_hash)

        tile_inv = self._create_module_shortcut_card(
            title="File Inventory & Audit",
            icon="folder",
            description="Recursively audit filesystem directory trees. Catalog comprehensive file metadata, identify hidden/system files, extract audio/video streams, and export structured Excel, CSV, and JSON audit sheets.",
            bullets=["Recursive directory scanning", "Media duration stream extraction", "XLSX, CSV & JSON deliverables"],
            btn_text="Open File Inventory",
            on_click=self._switch_to_inventory_mode
        )
        tiles_row.addWidget(tile_inv)

        tile_meta = self._create_module_shortcut_card(
            title="Forensic Metadata Explorer",
            icon="search",
            description="Deep forensic EXIF, XMP, and IPTC extraction powered by the official ExifTool executable. Captures raw JSON structures, verifies pre/post extraction hashes, and outputs HTML & Text reports.",
            bullets=[f"ExifTool {self.exiftool_ver} Integration", "Pre/Post Hash Verification", "Stand-alone HTML & TXT Reports"],
            btn_text="Open Metadata Explorer",
            on_click=self._switch_to_metadata_mode
        )
        tiles_row.addWidget(tile_meta)

        layout.addLayout(tiles_row)

        # 2. QUICK ACTIONS & ENVIRONMENT STATUS ROW (Original 2 Columns)
        qa_env_row = QHBoxLayout()
        qa_env_row.setSpacing(16)

        # Quick Actions Card
        qa_card = SurfaceCard(title="Quick Actions", icon="play")
        qa_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        qa_layout = QVBoxLayout()
        qa_layout.setSpacing(8)

        self.btn_qa_compare = QPushButton(" Compare Two Files Side-by-Side")
        self.btn_qa_compare.setIcon(LucideIcons.icon("compare", color="#2563eb" if theme == "light" else "#3b82f6", size=16))
        self.btn_qa_compare.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_qa_compare.setStyleSheet(self._action_btn_qss())
        self.btn_qa_compare.clicked.connect(self._open_compare_dialog)

        self.btn_qa_hash = QPushButton(" Quick Hash Single File")
        self.btn_qa_hash.setIcon(LucideIcons.icon("shield", color="#2563eb" if theme == "light" else "#3b82f6", size=16))
        self.btn_qa_hash.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_qa_hash.setStyleSheet(self._action_btn_qss())
        self.btn_qa_hash.clicked.connect(self._quick_action_hash_file)

        self.btn_qa_inv = QPushButton(" Quick Directory Audit")
        self.btn_qa_inv.setIcon(LucideIcons.icon("folder", color="#2563eb" if theme == "light" else "#3b82f6", size=16))
        self.btn_qa_inv.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_qa_inv.setStyleSheet(self._action_btn_qss())
        self.btn_qa_inv.clicked.connect(self._quick_action_inventory_folder)

        self.btn_qa_meta = QPushButton(" Quick Metadata Inspection")
        self.btn_qa_meta.setIcon(LucideIcons.icon("search", color="#2563eb" if theme == "light" else "#3b82f6", size=16))
        self.btn_qa_meta.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_qa_meta.setStyleSheet(self._action_btn_qss())
        self.btn_qa_meta.clicked.connect(self._quick_action_metadata_file)

        self.overview_action_buttons = [self.btn_qa_compare, self.btn_qa_hash, self.btn_qa_inv, self.btn_qa_meta]

        for btn in self.overview_action_buttons:
            btn.setFixedWidth(260)
            btn.setFixedHeight(34)
            row = QHBoxLayout()
            row.addWidget(btn)
            row.addStretch()
            qa_layout.addLayout(row)
        qa_card.main_layout.addLayout(qa_layout)
        qa_env_row.addWidget(qa_card, 1)

        # Workstation Environment Card
        env_card = SurfaceCard(title="Workstation Environment", icon="cpu")
        env_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        env_grid = QGridLayout()
        env_grid.setHorizontalSpacing(16)
        env_grid.setVerticalSpacing(10)

        # ExifTool status
        env_grid.addWidget(QLabel("Metadata Engine:"), 0, 0)
        self.overview_exif_status_lbl = QLabel(f"ExifTool {self.exiftool_ver} (Operational)" if self.exiftool_ok else "ExifTool Engine Not Detected")
        self.overview_exif_status_lbl.setStyleSheet("font-weight: 700; color: #287a45;" if theme == "light" else ("font-weight: 700; color: #80cc93;").replace("#80cc93", status_green(self)))
        env_grid.addWidget(self.overview_exif_status_lbl, 0, 1)

        # Forensic Integrity
        env_grid.addWidget(QLabel("Integrity Engine:"), 1, 0)
        lbl_integ = QLabel("Pre/Post SHA-256 Hash Guard Online")
        lbl_integ.setStyleSheet("font-weight: 600; color: #2563eb;" if theme == "light" else "font-weight: 600; color: #60a5fa;")
        env_grid.addWidget(lbl_integ, 1, 1)

        # Architecture & Python
        env_grid.addWidget(QLabel("Host Platform:"), 2, 0)
        lbl_plat = QLabel(f"Python {sys.version.split()[0]} • PySide6 • Windows")
        lbl_plat.setStyleSheet("color: #475569; font-size: 11px;" if theme == "light" else "color: #8e98b0; font-size: 11px;")
        env_grid.addWidget(lbl_plat, 2, 1)

        # Security mode
        env_grid.addWidget(QLabel("Security Mode:"), 3, 0)
        lbl_sec = QLabel("Air-Gapped Standalone (Zero Network Traffic)")
        lbl_sec.setStyleSheet("font-weight: 600; color: #287a45;" if theme == "light" else ("font-weight: 600; color: #80cc93;").replace("#80cc93", status_green(self)))
        env_grid.addWidget(lbl_sec, 3, 1)

        env_card.main_layout.addLayout(env_grid)
        qa_env_row.addWidget(env_card, 1)

        layout.addLayout(qa_env_row)

        # 3. SESSION ACTIVITY LOG CARD (Sleek, bounded height - no empty vertical void)
        self.activity_card = SurfaceCard(title="Session Activity Log", icon="activity")
        self.activity_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.activity_card.setMaximumHeight(250)
        self.activity_card.setMinimumHeight(150)
        act_layout = QVBoxLayout()
        act_layout.setContentsMargins(0, 0, 0, 0)
        act_layout.setSpacing(6)

        self.overview_empty_log_lbl = QLabel("No operations logged in current workstation session.\nActivities will appear here as you run Hash calculations, Directory audits, or Metadata extractions.")
        self.overview_empty_log_lbl.setAlignment(Qt.AlignCenter)
        empty_bg = "#f8fafc" if theme == "light" else "rgba(14, 18, 29, 0.4)"
        empty_color = "#64748b" if theme == "light" else "#8e98b0"
        empty_border = "#cbd5e1" if theme == "light" else "#232b43"
        self.overview_empty_log_lbl.setStyleSheet(f"color: {empty_color}; font-size: 11px; padding: 20px; background: {empty_bg}; border: 1px dashed {empty_border}; border-radius: 8px;")
        act_layout.addWidget(self.overview_empty_log_lbl)

        self.overview_activity_table = QTableWidget(0, 5)
        self.overview_activity_table.setHorizontalHeaderLabels(["Timestamp", "Module", "Target File / Folder", "Status", "Details"])
        self.overview_activity_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.overview_activity_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.overview_activity_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.overview_activity_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.overview_activity_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.overview_activity_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.overview_activity_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.overview_activity_table.verticalHeader().setVisible(False)
        self.overview_activity_table.setMaximumHeight(190)
        self.overview_activity_table.hide() # Hidden initially until operations happen
        act_layout.addWidget(self.overview_activity_table)

        self.activity_card.main_layout.addLayout(act_layout)
        layout.addWidget(self.activity_card)

        # Bottom stretch to absorb vertical space on widescreen and leave comfortable bottom margin
        layout.addStretch(1)

        scroll.setWidget(container)
        return scroll

    def _create_module_shortcut_card(self, title, icon, description, bullets, btn_text, on_click):
        """Create spacious module card with bounded height to avoid excessive vertical space on widescreen."""
        theme = getattr(self, "current_theme", "dark")
        card = SurfaceCard(title=title, icon=icon)
        card.setMaximumHeight(290)
        card.setMinimumHeight(240)

        clayout = QVBoxLayout()
        clayout.setSpacing(10)

        desc_lbl = QLabel(description)
        desc_lbl.setWordWrap(True)
        desc_color = "#475569" if theme == "light" else "#8e98b0"
        desc_lbl.setStyleSheet(f"font-size: 11px; line-height: 1.4; color: {desc_color};")
        if not hasattr(self, "overview_desc_labels"):
            self.overview_desc_labels = []
        self.overview_desc_labels.append(desc_lbl)
        clayout.addWidget(desc_lbl)

        bullets_box = QVBoxLayout()
        bullets_box.setSpacing(4)
        bullet_col = "#2563eb" if theme == "light" else "#60a5fa"
        txt_col = "#334155" if theme == "light" else "#cbd5e1"
        for b in bullets:
            row = QHBoxLayout()
            row.setSpacing(6)
            dot = QLabel("•")
            dot.setStyleSheet(f"color: {bullet_col}; font-size: 13px; font-weight: 900;")
            blbl = QLabel(b)
            blbl.setStyleSheet(f"font-size: 11px; color: {txt_col}; font-weight: 500;")
            if not hasattr(self, "overview_bullet_labels"):
                self.overview_bullet_labels = []
            self.overview_bullet_labels.append((dot, blbl))
            row.addWidget(dot)
            row.addWidget(blbl)
            row.addStretch()
            bullets_box.addLayout(row)
        clayout.addLayout(bullets_box)

        clayout.addStretch()

        clean_btn_text = btn_text.replace("→", "").replace("->", "").strip()
        btn = QPushButton(clean_btn_text)
        btn.setCursor(QCursor(Qt.PointingHandCursor))
        btn.setMinimumWidth(180)
        btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        btn.setFixedHeight(34)
        self._style_button_primary(btn, theme)
        btn.clicked.connect(on_click)
        if not hasattr(self, "overview_shortcut_buttons"):
            self.overview_shortcut_buttons = []
        self.overview_shortcut_buttons.append(btn)

        btn_row = QHBoxLayout()
        btn_row.addWidget(btn)
        btn_row.addStretch()
        clayout.addLayout(btn_row)

        card.main_layout.addLayout(clayout)
        return card

    def _action_btn_qss(self):
        theme = getattr(self, "current_theme", "dark")
        if theme == "light":
            return """
                QPushButton {
                    background-color: #f8fafc;
                    border: 1px solid #cbd5e1;
                    border-radius: 8px;
                    color: #0f172a;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                    text-align: left;
                }
                QPushButton:hover {
                    background-color: #eff6ff;
                    border-color: #2563eb;
                    color: #1d4ed8;
                }
            """
        else:
            return """
                QPushButton {
                    background-color: #121727;
                    border: 1px solid #232b43;
                    border-radius: 8px;
                    color: #f1f5f9;
                    padding: 8px 14px;
                    font-size: 12px;
                    font-weight: 600;
                    text-align: left;
                }
                QPushButton:hover {
                    background-color: #1a2238;
                    border-color: #3b82f6;
                    color: #ffffff;
                }
            """

    def _quick_action_hash_file(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files for Cryptographic Hashing")
        if files:
            for f in files:
                self._add_file_to_hash_items(f)
            self._refresh_hash_table()
            self._switch_to_hash_mode()

    def _quick_action_inventory_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Inventory")
        if folder:
            self.inv_src_edit.setText(folder)
            self._switch_to_inventory_mode()

    def _quick_action_metadata_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Evidence File for Metadata", "", "All Files (*.*)")
        if path:
            self.meta_rb_file.setChecked(True)
            self.meta_src_edit.setText(path)
            self._on_meta_path_dropped(path)
            self._switch_to_metadata_mode()

    def _log_session_activity(self, module: str, target: str, status: str, details: str):
        """Append an entry to the real session activity log and update the Overview table."""
        record = {
            "time": datetime.datetime.now().strftime("%H:%M:%S"),
            "module": module,
            "target": os.path.basename(target) if target else "N/A",
            "status": status,
            "details": details
        }
        self.session_activity_log.insert(0, record)  # Newest first

        if hasattr(self, "overview_activity_table") and self.overview_activity_table:
            self._refresh_overview_activity_table()

    def _refresh_overview_activity_table(self):
        if not hasattr(self, "overview_activity_table") or not self.overview_activity_table:
            return

        count = len(self.session_activity_log)
        if count == 0:
            if hasattr(self, "overview_empty_log_lbl"):
                self.overview_empty_log_lbl.show()
            self.overview_activity_table.hide()
            return

        if hasattr(self, "overview_empty_log_lbl"):
            self.overview_empty_log_lbl.hide()
        self.overview_activity_table.show()

        self.overview_activity_table.setRowCount(min(count, 50))
        for row, item in enumerate(self.session_activity_log[:50]):
            time_item = QTableWidgetItem(item["time"])
            time_item.setTextAlignment(Qt.AlignCenter)

            mod_item = QTableWidgetItem(item["module"])

            target_item = QTableWidgetItem(item["target"])
            target_item.setToolTip(item["target"])

            status_item = QTableWidgetItem(item["status"])
            status_item.setTextAlignment(Qt.AlignCenter)
            theme = getattr(self, "current_theme", "dark")
            if item["status"] in ("Completed", "Verified"):
                status_item.setForeground(QColor("#287a45" if theme == "light" else ("#80cc93").replace("#80cc93", status_green(self))))
            elif item["status"] in ("Alert", "Error"):
                status_item.setForeground(QColor("#dc2626" if theme == "light" else "#f87171"))
            else:
                status_item.setForeground(QColor("#2563eb" if theme == "light" else "#60a5fa"))

            details_item = QTableWidgetItem(item["details"])

            self.overview_activity_table.setItem(row, 0, time_item)
            self.overview_activity_table.setItem(row, 1, mod_item)
            self.overview_activity_table.setItem(row, 2, target_item)
            self.overview_activity_table.setItem(row, 3, status_item)
            self.overview_activity_table.setItem(row, 4, details_item)

    # ==========================================================================
    # WORKSTATION SETTINGS PAGE
    # ==========================================================================
    def _create_settings_page(self):
        """
        Constructs the Workstation Settings screen:
        1. ExifTool Engine Executable Path & Detection
        2. Cryptographic Hash Preferences (Default Suites & Chunk Buffer)
        3. Report Deliverables & Default Output Folder
        4. Interface & Persistence Preferences (Theme & Startup Module)
        5. Action buttons to save or restore defaults
        """
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(6, 6, 6, 20)
        layout.setSpacing(14)

        # 1. EXIFTOOL ENGINE CONFIGURATION CARD
        exif_card = SurfaceCard(title="ExifTool Metadata Engine", icon="terminal")
        exif_layout = QVBoxLayout()
        exif_layout.setSpacing(10)

        exif_desc = QLabel("The Metadata Explorer module relies on the official ExifTool standalone executable for EXIF, XMP, IPTC, and structural metadata parsing. Configure the binary location below.")
        exif_desc.setWordWrap(True)
        exif_desc.setStyleSheet("font-size: 12px; color: #8e98b0;")
        exif_layout.addWidget(exif_desc)

        path_row = QHBoxLayout()
        self.settings_exif_path_edit = QLineEdit()
        self.settings_exif_path_edit.setPlaceholderText("Auto-detect default or browse to exiftool.exe...")
        if self.custom_exiftool_path:
            self.settings_exif_path_edit.setText(self.custom_exiftool_path)
        elif self.exiftool_path:
            self.settings_exif_path_edit.setText(self.exiftool_path)

        btn_browse_exif = QPushButton("Browse")
        btn_browse_exif.setCursor(QCursor(Qt.PointingHandCursor))
        btn_browse_exif.setStyleSheet(self._settings_btn_qss())
        btn_browse_exif.clicked.connect(self._browse_exiftool_engine)

        btn_auto_exif = QPushButton("Auto-Detect")
        btn_auto_exif.setCursor(QCursor(Qt.PointingHandCursor))
        btn_auto_exif.setStyleSheet(self._settings_btn_qss())
        btn_auto_exif.clicked.connect(self._autodetect_exiftool_engine)

        btn_test_exif = QPushButton("Test Engine")
        btn_test_exif.setCursor(QCursor(Qt.PointingHandCursor))
        btn_test_exif.setStyleSheet(self._settings_btn_primary_qss())
        btn_test_exif.clicked.connect(self._test_exiftool_engine)

        self.settings_buttons = [btn_browse_exif, btn_auto_exif]

        path_row.addWidget(self.settings_exif_path_edit)
        path_row.addWidget(btn_browse_exif)
        path_row.addWidget(btn_auto_exif)
        path_row.addWidget(btn_test_exif)
        exif_layout.addLayout(path_row)

        self.settings_exif_status_lbl = QLabel()
        self._update_settings_exif_status_display()
        exif_layout.addWidget(self.settings_exif_status_lbl)

        exif_card.main_layout.addLayout(exif_layout)
        layout.addWidget(exif_card)

        # 2. CRYPTOGRAPHIC HASH PREFERENCES CARD
        hash_card = SurfaceCard(title="Cryptographic Hash Preferences", icon="shield")
        h_layout = QVBoxLayout()
        h_layout.setSpacing(10)

        h_desc = QLabel("Configure standard hashing algorithms calculated during file audits and integrity checks.")
        h_desc.setStyleSheet("font-size: 12px; color: #8e98b0;")
        h_layout.addWidget(h_desc)

        algos_row = QHBoxLayout()
        self.settings_cb_sha256 = QCheckBox("SHA-256 (Forensic Standard - Required)")
        self.settings_cb_sha256.setChecked(True)
        self.settings_cb_sha256.setEnabled(False) # Always mandatory
        self.settings_cb_sha512 = QCheckBox("SHA-512")
        self.settings_cb_sha512.setChecked(self.settings_store.value("hash_enabled_SHA-512", True, type=bool))
        self.settings_cb_md5 = QCheckBox("MD5")
        self.settings_cb_md5.setChecked(self.settings_store.value("hash_enabled_MD5", True, type=bool))
        self.settings_cb_sha1 = QCheckBox("SHA-1")
        self.settings_cb_sha1.setChecked(self.settings_store.value("hash_enabled_SHA-1", True, type=bool))

        algos_row.addWidget(self.settings_cb_sha256)
        algos_row.addWidget(self.settings_cb_sha512)
        algos_row.addWidget(self.settings_cb_md5)
        algos_row.addWidget(self.settings_cb_sha1)
        algos_row.addStretch()
        h_layout.addLayout(algos_row)

        chunk_row = QHBoxLayout()
        chunk_lbl = QLabel("Hash Stream Buffer Chunk Size:")
        chunk_lbl.setStyleSheet("font-size: 12px; color: #8e98b0;")
        self.settings_chunk_combo = QComboBox()
        self.settings_chunk_combo.addItems(["64 KB", "128 KB (Recommended)", "256 KB", "1 MB", "4 MB (High-Throughput NVMe)"])
        saved_chunk = self.settings_store.value("chunk_size_index", 1, type=int)
        self.settings_chunk_combo.setCurrentIndex(saved_chunk)
        chunk_row.addWidget(chunk_lbl)
        chunk_row.addWidget(self.settings_chunk_combo)
        chunk_row.addStretch()
        h_layout.addLayout(chunk_row)

        hash_card.main_layout.addLayout(h_layout)
        layout.addWidget(hash_card)

        # 3. REPORT OUTPUT & STORAGE PREFERENCES CARD
        out_card = SurfaceCard(title="Report Deliverables & Output Directory", icon="database")
        out_layout = QVBoxLayout()
        out_layout.setSpacing(10)

        out_desc = QLabel("Set default output directory for generated Excel audit spreadsheets, CSV tables, and ExifTool reports.")
        out_desc.setStyleSheet("font-size: 12px; color: #8e98b0;")
        out_layout.addWidget(out_desc)

        out_row = QHBoxLayout()
        self.settings_output_edit = QLineEdit()
        default_reports_dir = str(Path.home() / "Downloads" / "FileInventory_Reports")
        saved_out = self.settings_store.value("default_output_dir", default_reports_dir)
        self.settings_output_edit.setText(saved_out)
        self.settings_output_edit.setPlaceholderText("Select default reports directory...")

        btn_browse_out = QPushButton("Browse")
        btn_browse_out.setCursor(QCursor(Qt.PointingHandCursor))
        btn_browse_out.setStyleSheet(self._settings_btn_qss())
        btn_browse_out.clicked.connect(self._browse_settings_output_folder)
        self.settings_buttons.append(btn_browse_out)

        out_row.addWidget(self.settings_output_edit)
        out_row.addWidget(btn_browse_out)
        out_layout.addLayout(out_row)

        size_row = QHBoxLayout()
        size_row.addWidget(QLabel("File Inventory report sizes:"))
        self.settings_inventory_size_combo = QComboBox()
        for title, value in (("Automatic (B, KB, MB, GB...)", "auto"),
                             ("Kilobytes (KB)", "kb"), ("Megabytes (MB)", "mb")):
            self.settings_inventory_size_combo.addItem(title, value)
        saved_size_unit = self.settings_store.value("inventory_size_unit", "auto")
        size_index = self.settings_inventory_size_combo.findData(saved_size_unit)
        self.settings_inventory_size_combo.setCurrentIndex(max(0, size_index))
        self.settings_inventory_size_combo.setToolTip(
            "Size display in File Inventory Excel/HTML reports. 1 KB = 1024 bytes; 1 MB = 1024 KB.")
        size_row.addWidget(self.settings_inventory_size_combo)
        size_row.addStretch()
        out_layout.addLayout(size_row)

        out_card.main_layout.addLayout(out_layout)
        layout.addWidget(out_card)

        # 4. WORKSTATION INTERFACE & STARTUP CARD
        ui_card = SurfaceCard(title="Workstation Interface & Startup", icon="settings")
        ui_layout = QVBoxLayout()
        ui_layout.setSpacing(12)

        theme_lbl = QLabel("Color Theme:")
        theme_lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #8e98b0;")
        ui_layout.addWidget(theme_lbl)

        theme_row = QHBoxLayout()
        self.settings_rb_theme_dark = QRadioButton("Forensic Charcoal (Dark Mode)")
        self.settings_rb_theme_light = QRadioButton("Clean Workstation (Light Mode)")
        if self.current_theme == "light":
            self.settings_rb_theme_light.setChecked(True)
        else:
            self.settings_rb_theme_dark.setChecked(True)
        theme_row.addWidget(self.settings_rb_theme_dark)
        theme_row.addWidget(self.settings_rb_theme_light)
        theme_row.addStretch()
        ui_layout.addLayout(theme_row)

        startup_lbl = QLabel("Default Startup View:")
        startup_lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #8e98b0;")
        ui_layout.addWidget(startup_lbl)

        startup_row = QHBoxLayout()
        self.settings_startup_combo = QComboBox()
        self.settings_startup_combo.addItems([
            "Workstation Overview",
            "Hash Verification",
            "File Inventory",
            "Metadata Explorer"
        ])
        saved_startup = self.settings_store.value("startup_tab", 0, type=int)
        self.settings_startup_combo.setCurrentIndex(saved_startup if saved_startup in (0, 1, 2, 3) else 0)
        startup_row.addWidget(self.settings_startup_combo)
        startup_row.addStretch()
        ui_layout.addLayout(startup_row)

        ui_card.main_layout.addLayout(ui_layout)
        layout.addWidget(ui_card)

        # 5. SAVE & RESET BUTTONS
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        btn_save = QPushButton("Save Settings")
        btn_save.setCursor(QCursor(Qt.PointingHandCursor))
        btn_save.setStyleSheet(self._settings_btn_primary_qss())
        btn_save.clicked.connect(self._save_workstation_settings)

        btn_reset = QPushButton("Reset to Defaults")
        btn_reset.setCursor(QCursor(Qt.PointingHandCursor))
        btn_reset.setStyleSheet(self._settings_btn_qss())
        btn_reset.clicked.connect(self._reset_workstation_settings)
        self.settings_buttons.append(btn_reset)

        self.settings_feedback_lbl = QLabel("")
        self.settings_feedback_lbl.setStyleSheet(("font-size: 11px; font-weight: 600; color: #80cc93;").replace("#80cc93", status_green(self)))

        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_reset)
        self.settings_about_btn = QPushButton("About FINDEX")
        self.settings_about_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.settings_about_btn.setStyleSheet(self._settings_btn_qss())
        self.settings_about_btn.clicked.connect(self._show_about_findex)
        self.settings_buttons.append(self.settings_about_btn)
        btn_row.addWidget(self.settings_about_btn)
        btn_row.addWidget(self.settings_feedback_lbl)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        layout.addStretch()
        scroll.setWidget(container)
        return scroll

    def _show_about_findex(self):
        AboutFindexDialog(self).exec()

    def _settings_btn_qss(self):
        theme = getattr(self, "current_theme", "dark")
        if theme == "light":
            return """
                QPushButton {
                    background-color: #f1f5f9;
                    border: 1px solid #cbd5e1;
                    border-radius: 8px;
                    color: #0f172a;
                    padding: 8px 16px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #e2e8f0;
                    border-color: #2563eb;
                }
            """
        else:
            return """
                QPushButton {
                    background-color: #1e243b;
                    border: 1px solid #2e3654;
                    border-radius: 8px;
                    color: #f1f5f9;
                    padding: 8px 16px;
                    font-size: 12px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #28304e;
                    border-color: #3b82f6;
                }
            """

    def _settings_btn_primary_qss(self):
        return """
            QPushButton {
                background-color: #3b82f6;
                border: none;
                border-radius: 8px;
                color: #ffffff;
                padding: 8px 18px;
                font-size: 12px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #60a5fa;
            }
        """

    def _update_settings_exif_status_display(self):
        if not hasattr(self, "settings_exif_status_lbl") or not self.settings_exif_status_lbl:
            return
        if self.exiftool_ok:
            self.settings_exif_status_lbl.setText(f"Operational: ExifTool v{self.exiftool_ver} detected at '{self.exiftool_path}'")
            self.settings_exif_status_lbl.setStyleSheet(("font-size: 11px; font-weight: 600; color: #80cc93; background: rgba(128, 204, 147, 0.12); border: 1px solid rgba(128, 204, 147, 0.25); border-radius: 6px; padding: 6px 10px;").replace("#80cc93", status_green(self)))
        else:
            self.settings_exif_status_lbl.setText("ExifTool executable not found. Please browse to a valid exiftool.exe binary.")
            self.settings_exif_status_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #f87171; background: rgba(239, 68, 68, 0.12); border: 1px solid rgba(239, 68, 68, 0.25); border-radius: 6px; padding: 6px 10px;")

    def _update_sidebar_engine_display(self):
        if hasattr(self, "sidebar_engine_pill") and self.sidebar_engine_pill:
            if self.exiftool_ok:
                self.sidebar_engine_pill.setText(f"ExifTool v{self.exiftool_ver}")
            else:
                self.sidebar_engine_pill.setText("ExifTool Not Found")
        if hasattr(self, "overview_exif_status_lbl") and self.overview_exif_status_lbl:
            self.overview_exif_status_lbl.setText(f"ExifTool {self.exiftool_ver} (Operational)" if self.exiftool_ok else "ExifTool Engine Not Detected")
            self.overview_exif_status_lbl.setStyleSheet(("font-weight: 700; color: #80cc93;").replace("#80cc93", status_green(self)) if self.exiftool_ok else "font-weight: 700; color: #f87171;")

    def _browse_exiftool_engine(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select ExifTool Executable", "", "Executables (*.exe);;All Files (*.*)")
        if p:
            self.settings_exif_path_edit.setText(p)
            self._test_exiftool_engine()

    def _autodetect_exiftool_engine(self):
        ver, ok, path = ExifToolEngine.get_version(None)
        if ok and path:
            self.settings_exif_path_edit.setText(path)
            self.exiftool_ver = ver
            self.exiftool_ok = ok
            self.exiftool_path = path
            self._update_settings_exif_status_display()
            self._update_sidebar_engine_display()
            if hasattr(self, "_update_engine_status_pill"):
                self._update_engine_status_pill()
        else:
            QMessageBox.warning(self, "ExifTool Not Detected", "Could not locate ExifTool automatically in external/exiftool/ or system PATH.")

    def _test_exiftool_engine(self):
        p = self.settings_exif_path_edit.text().strip()
        ver, ok, path = ExifToolEngine.get_version(p if p else None)
        self.exiftool_ver = ver
        self.exiftool_ok = ok
        self.exiftool_path = path
        self._update_settings_exif_status_display()
        self._update_sidebar_engine_display()
        if hasattr(self, "_update_engine_status_pill"):
            self._update_engine_status_pill()
        if ok:
            QMessageBox.information(self, "Engine Test Successful", f"ExifTool Engine is operational!\nDetected Version: {ver}\nPath: {path}")
        else:
            QMessageBox.critical(self, "Engine Test Failed", f"Failed to execute ExifTool at '{p}'.\nPlease verify the executable exists and has execute permissions.")

    def _browse_settings_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Default Reports Output Directory")
        if folder:
            self.settings_output_edit.setText(folder)

    def _save_workstation_settings(self):
        # Save ExifTool path
        custom_exif = self.settings_exif_path_edit.text().strip()
        if custom_exif and os.path.exists(custom_exif):
            self.settings_store.setValue("custom_exiftool_path", custom_exif)
            self.custom_exiftool_path = custom_exif
        else:
            self.settings_store.remove("custom_exiftool_path")
            self.custom_exiftool_path = None

        # Save buffer chunk size
        self.settings_store.setValue("chunk_size_index", self.settings_chunk_combo.currentIndex())
        for algorithm, widget in (("MD5", self.settings_cb_md5), ("SHA-1", self.settings_cb_sha1), ("SHA-512", self.settings_cb_sha512)):
            self.settings_store.setValue("hash_enabled_" + algorithm, widget.isChecked())
        configure_hash_preferences(self.settings_store)

        # Save default output folder
        self.settings_store.setValue("inventory_size_unit", self.settings_inventory_size_combo.currentData())
        out_dir = self.settings_output_edit.text().strip()
        if out_dir:
            self.settings_store.setValue("default_output_dir", out_dir)
            if hasattr(self, "inv_out_edit"):
                if os.path.isdir(self.inv_src_edit.text().strip()):
                    self._auto_populate_output_paths(self.inv_src_edit.text())
                else:
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    self.inv_out_edit.setText(str(Path(out_dir) / f"File_Inventory_{timestamp}"))
            if hasattr(self, "meta_out_dir_edit"):
                self.meta_out_dir_edit.setText(out_dir)

        # Save auto open folder
        self.settings_store.remove("auto_open_folder")

        # Save theme preference
        chosen_theme = "light" if self.settings_rb_theme_light.isChecked() else "dark"
        self.settings_store.setValue("theme", chosen_theme)
        if chosen_theme != self.current_theme:
            self._apply_global_theme(chosen_theme)

        # Save startup view
        self.settings_store.setValue("startup_tab", self.settings_startup_combo.currentIndex())

        self.settings_feedback_lbl.setText("Settings successfully saved and applied.")
        QTimer.singleShot(4000, lambda: self.settings_feedback_lbl.setText(""))

    def _reset_workstation_settings(self):
        self.settings_store.clear()
        configure_hash_preferences(self.settings_store)
        for widget in (self.settings_cb_md5, self.settings_cb_sha1, self.settings_cb_sha512):
            widget.setChecked(True)
        self.settings_exif_path_edit.clear()
        self.custom_exiftool_path = None
        self._autodetect_exiftool_engine()
        default_reports_dir = str(Path.home() / "Downloads" / "FileInventory_Reports")
        self.settings_output_edit.setText(default_reports_dir)
        self.settings_chunk_combo.setCurrentIndex(1)
        self.settings_inventory_size_combo.setCurrentIndex(0)
        self.settings_startup_combo.setCurrentIndex(0)
        self.settings_rb_theme_dark.setChecked(True)
        self._apply_global_theme("dark")
        self.settings_feedback_lbl.setText("Workstation settings reset to defaults.")
        QTimer.singleShot(4000, lambda: self.settings_feedback_lbl.setText(""))

    # ==========================================================================
    # PAGE 1: FILE INVENTORY & AUDIT MODE
    # ==========================================================================
    def _create_inventory_page(self):
        page = QWidget()
        page_layout = QHBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(16)

        # ── LEFT PANEL: CONFIGURATION (Expands to majority of width) ──
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 10, 0)
        left_layout.setSpacing(12)

        # 1. Source Folder Card
        src_card = SurfaceCard(title="Source Folder", icon="folder")
        src_row = QHBoxLayout()
        self.inv_src_edit = QLineEdit()
        self.inv_src_edit.setPlaceholderText("Choose parent folder to scan...")
        self.inv_src_edit.textChanged.connect(self._auto_populate_output_paths)
        self.inv_btn_src_browse = QPushButton("Browse")
        self.inv_btn_src_browse.setCursor(QCursor(Qt.PointingHandCursor))
        self.inv_btn_src_browse.setFixedWidth(85)
        self._style_button_secondary(self.inv_btn_src_browse, getattr(self, "current_theme", "dark"))
        self.inv_btn_src_browse.clicked.connect(self._browse_source_folder)
        src_row.addWidget(self.inv_src_edit)
        src_row.addWidget(self.inv_btn_src_browse)
        src_card.main_layout.addLayout(src_row)
        left_layout.addWidget(src_card)

        # 2. Output Format & Path Card
        out_card = SurfaceCard(title="Output Format & Path", icon="file_text")
        fmt_row = QHBoxLayout()
        self.inv_fmt_lbl = QLabel("Format:")
        self.inv_fmt_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #475569;" if getattr(self, "current_theme", "dark") == "light" else "font-size: 11px; font-weight: 700; color: #8e98b0;")
        fmt_lbl = self.inv_fmt_lbl
        fmt_row.addWidget(fmt_lbl)

        self.rb_both = QRadioButton("Excel + HTML")
        self.rb_excel = QRadioButton("Excel (.xlsx)")
        self.rb_html = QRadioButton("HTML (.html)")
        self.rb_both.setChecked(True)
        fmt_group = QButtonGroup(self)
        fmt_group.addButton(self.rb_both)
        fmt_group.addButton(self.rb_excel)
        fmt_group.addButton(self.rb_html)
        fmt_row.addWidget(self.rb_both)
        fmt_row.addWidget(self.rb_excel)
        fmt_row.addWidget(self.rb_html)
        fmt_row.addStretch()
        out_card.main_layout.addLayout(fmt_row)

        out_row = QHBoxLayout()
        self.inv_out_path_lbl = QLabel("Save To:")
        self.inv_out_path_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #475569;" if getattr(self, "current_theme", "dark") == "light" else "font-size: 11px; font-weight: 700; color: #8e98b0;")
        out_path_lbl = self.inv_out_path_lbl
        out_row.addWidget(out_path_lbl)

        self.inv_out_edit = QLineEdit()
        self.inv_out_edit.setPlaceholderText("Default location (auto-filled on folder selection)...")
        self.inv_btn_out_browse = QPushButton("Browse")
        self.inv_btn_out_browse.setCursor(QCursor(Qt.PointingHandCursor))
        self.inv_btn_out_browse.setFixedWidth(85)
        self._style_button_secondary(self.inv_btn_out_browse, getattr(self, "current_theme", "dark"))
        self.inv_btn_out_browse.clicked.connect(self._browse_output_file)
        out_row.addWidget(self.inv_out_edit)
        out_row.addWidget(self.inv_btn_out_browse)
        out_card.main_layout.addLayout(out_row)
        left_layout.addWidget(out_card)

        # 3. File Types Card (3 Columns across the wide view)
        filter_card = SurfaceCard(title="File Types", icon="layers")
        grid_filters = QGridLayout()
        grid_filters.setSpacing(10)
        self.cb_audio = QCheckBox("Audio (.mp3, .wav, .flac, etc.)")
        self.cb_video = QCheckBox("Video (.mp4, .mkv, .mov, etc.)")
        self.cb_images = QCheckBox("Images (.jpg, .png, etc.)")
        self.cb_docs = QCheckBox("Documents (.pdf, .docx, etc.)")
        self.cb_archives = QCheckBox("Archives (.zip, .rar, etc.)")
        self.cb_other = QCheckBox("Other Extensions")

        self.cb_audio.setChecked(True)
        self.cb_video.setChecked(True)
        self.cb_images.setChecked(False)
        self.cb_docs.setChecked(False)
        self.cb_archives.setChecked(False)
        self.cb_other.setChecked(False)

        grid_filters.addWidget(self.cb_audio, 0, 0)
        grid_filters.addWidget(self.cb_video, 0, 1)
        grid_filters.addWidget(self.cb_images, 0, 2)
        grid_filters.addWidget(self.cb_docs, 1, 0)
        grid_filters.addWidget(self.cb_archives, 1, 1)
        grid_filters.addWidget(self.cb_other, 1, 2)
        filter_card.main_layout.addLayout(grid_filters)
        left_layout.addWidget(filter_card)

        # 4. Columns & Preset Card (4 Columns in wide view)
        col_card = SurfaceCard(title="Columns & Presets", icon="database")
        preset_row = QHBoxLayout()
        preset_lbl = QLabel("Preset:")
        preset_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #8e98b0;")
        preset_row.addWidget(preset_lbl)

        self.preset_combo = QComboBox()
        self.preset_combo.addItem("Media Details (Default)", "media")
        self.preset_combo.addItem("Forensic (With Hashes)", "forensic")
        self.preset_combo.addItem("All Columns", "complete")
        self.preset_combo.addItem("Compact Minimal", "minimal")
        self.preset_combo.currentIndexChanged.connect(self._on_preset_changed)
        preset_row.addWidget(self.preset_combo)
        preset_row.addStretch()
        col_card.main_layout.addLayout(preset_row)

        col_grid = QGridLayout()
        col_grid.setSpacing(8)
        self.col_checkboxes = {}

        cols_display = [
            ("si", "SI No.", True, True),
            ("name", "File Name", True, True),
            ("size", "File Size", True, False),
            ("format", "Format", True, False),
            ("duration", "Duration", True, False),
            ("modified", "Modified Date", True, False),
            ("created", "Created Date", False, False),
            ("path", "Full Path", False, False),
            ("md5", "MD5", False, False),
            ("sha1", "SHA-1", False, False),
            ("sha256", "SHA-256", False, False),
        ]

        # Arrange across 4 columns for a clean modern dashboard look
        for idx, (cid, cname, cdef, clocked) in enumerate(cols_display):

            r = idx // 4

            c = idx % 4
            cb = QCheckBox(cname)
            cb.setChecked(cdef)
            if clocked:
                cb.setEnabled(False)
            self.col_checkboxes[cid] = cb
            col_grid.addWidget(cb, r, c)
        col_card.main_layout.addLayout(col_grid)
        left_layout.addWidget(col_card)

        # 5. Options Card (SI Numbering + Empty Folder Filter)
        opt_card = SurfaceCard(title="Options", icon="settings")
        opt_layout = QVBoxLayout()
        opt_layout.setSpacing(8)

        si_row = QHBoxLayout()
        si_lbl = QLabel("SI Numbering:")
        si_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #8e98b0;")
        si_row.addWidget(si_lbl)

        self.rb_si_folder = QRadioButton("Per folder (start from 1)")
        self.rb_si_continuous = QRadioButton("Continuous numbering")
        self.rb_si_folder.setChecked(True)
        si_group = QButtonGroup(self)
        si_group.addButton(self.rb_si_folder)
        si_group.addButton(self.rb_si_continuous)
        si_row.addWidget(self.rb_si_folder)
        si_row.addWidget(self.rb_si_continuous)
        si_row.addStretch()
        opt_layout.addLayout(si_row)

        self.cb_skip_empty_folders = QCheckBox("Exclude folders containing no matching files")
        self.cb_skip_empty_folders.setChecked(True)
        opt_layout.addWidget(self.cb_skip_empty_folders)

        opt_card.main_layout.addLayout(opt_layout)
        left_layout.addWidget(opt_card)
        left_layout.addStretch()

        scroll.setWidget(left_container)
        scroll.setMinimumWidth(360)
        page_layout.addWidget(scroll, 58)

        # ── RIGHT PANEL: STATISTICS & STATUS (Flexible Scrollable Container) ──
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        right_scroll.setFrameShape(QFrame.NoFrame)
        right_scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        right_container = QWidget()
        right_container.setMinimumWidth(280)
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 8, 0)
        right_layout.setSpacing(12)

        # 4 KPI Stat Cards in 2x2 Grid
        kpi_grid = QGridLayout()
        kpi_grid.setSpacing(8)

        self.stat_files = ModernStatCard("file_text", "Files", "0", "")
        self.stat_size = ModernStatCard("database", "Total Size", "0 B", "")
        self.stat_folders = ModernStatCard("folder", "Folders", "0", "")
        self.stat_time = ModernStatCard("clock", "Time", "0.0s", "")

        kpi_grid.addWidget(self.stat_files, 0, 0)
        kpi_grid.addWidget(self.stat_size, 0, 1)
        kpi_grid.addWidget(self.stat_folders, 1, 0)
        kpi_grid.addWidget(self.stat_time, 1, 1)
        right_layout.addLayout(kpi_grid)

        # Progress / Status Card with animated Processing Spinner
        self.progress_card = SurfaceCard(title="Status", icon="activity")

        status_row = QHBoxLayout()
        status_row.setSpacing(10)
        self.inv_spinner = ProcessingSpinner(size=24)
        status_row.addWidget(self.inv_spinner)

        status_text_layout = QVBoxLayout()
        status_text_layout.setSpacing(2)
        self.inv_status_lbl = QLabel("Ready")
        self.inv_status_lbl.setStyleSheet("font-size: 13px; font-weight: 700; color: #f1f5f9;")
        self.inv_ticker_lbl = QLabel("—")
        self.inv_ticker_lbl.setStyleSheet("font-family: monospace; font-size: 11px; color: #8e98b0;")
        status_text_layout.addWidget(self.inv_status_lbl)
        status_text_layout.addWidget(self.inv_ticker_lbl)
        status_row.addLayout(status_text_layout, 1)
        self.progress_card.main_layout.addLayout(status_row)

        self.inv_progress_bar = QProgressBar()
        self.inv_progress_bar.setRange(0, 100)
        self.inv_progress_bar.setValue(0)
        self.progress_card.main_layout.addWidget(self.inv_progress_bar)
        right_layout.addWidget(self.progress_card)

        # Action Execution Card (Pinned directly below Status & Progress)
        action_card = SurfaceCard()
        act_row = QHBoxLayout()
        act_row.setSpacing(12)
        act_row.addStretch()
        self.btn_run_inventory = QPushButton("Run Inventory")
        self.btn_run_inventory.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_run_inventory.setFixedWidth(170)
        self.btn_run_inventory.setFixedHeight(36)
        self.btn_run_inventory.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: #ffffff;
                font-size: 12px;
                font-weight: 700;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #60a5fa; }
            QPushButton:pressed { background-color: #2563eb; }
        """)
        self.btn_run_inventory.clicked.connect(self._start_inventory_scan)

        self.btn_cancel_inventory = QPushButton("Cancel")
        self.btn_cancel_inventory.setEnabled(False)
        self.btn_cancel_inventory.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_cancel_inventory.setFixedWidth(90)
        self.btn_cancel_inventory.setFixedHeight(36)
        self.btn_cancel_inventory.setStyleSheet("""
            QPushButton {
                background-color: #ef4444;
                color: #ffffff;
                font-size: 12px;
                font-weight: 700;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #f87171; }
            QPushButton:disabled { background-color: #1e243b; color: #55607c; }
        """)
        self.btn_cancel_inventory.clicked.connect(self._cancel_inventory_scan)

        act_row.addWidget(self.btn_run_inventory)
        act_row.addWidget(self.btn_cancel_inventory)
        act_row.addStretch()
        action_card.main_layout.addLayout(act_row)
        right_layout.addWidget(action_card)

        # Output Deliverables Card
        self.results_card = SurfaceCard(title="Output Files", icon="external_link")
        self.results_desc = QLabel("Generated files will appear here.")
        self.results_desc.setStyleSheet("font-size: 11px; color: #8e98b0;")
        self.results_card.main_layout.addWidget(self.results_desc)

        self.results_btn_layout = QVBoxLayout()
        self.results_btn_layout.setSpacing(8)
        self.results_card.main_layout.addLayout(self.results_btn_layout)
        right_layout.addWidget(self.results_card)

        right_layout.addStretch()
        right_scroll.setWidget(right_container)
        page_layout.addWidget(right_scroll, 42)

        return page

    def _browse_source_folder(self):
        f = QFileDialog.getExistingDirectory(self, "Select Parent Folder to Inventory")
        if f:
            self.inv_src_edit.setText(f)

    def _auto_populate_output_paths(self, src_text):
        if not src_text or not os.path.isdir(src_text):
            return
        p = Path(src_text)
        fname = p.name or "File_Inventory"
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = Path(self.settings_store.value("default_output_dir", str(p.parent)) or p.parent)
        out_base = str(output_dir / f"{fname}_Inventory_{ts}")
        self.inv_out_edit.setText(out_base)

    def _browse_output_file(self):
        f, _ = QFileDialog.getSaveFileName(self, "Specify Base Output Path and Filename",
                                         options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            base_no_ext = str(Path(f).with_suffix(""))
            self.inv_out_edit.setText(base_no_ext)

    def _on_preset_changed(self, idx):
        preset_key = self.preset_combo.currentData()
        self._apply_column_preset(preset_key)

    def _apply_column_preset(self, preset_key):
        for cid, cb in self.col_checkboxes.items():
            if cid in ("si", "name"):
                continue
            cb.setChecked(False)

        if preset_key == "media":
            for cid in ("size", "format", "duration", "modified"):
                if cid in self.col_checkboxes:
                    self.col_checkboxes[cid].setChecked(True)
        elif preset_key == "forensic":
            for cid in ("size", "format", "duration", "modified", "path", "md5", "sha1", "sha256"):
                if cid in self.col_checkboxes:
                    self.col_checkboxes[cid].setChecked(True)
        elif preset_key == "complete":
            for cb in self.col_checkboxes.values():
                cb.setChecked(True)
        elif preset_key == "minimal":
            for cid in ("size", "modified"):
                if cid in self.col_checkboxes:
                    self.col_checkboxes[cid].setChecked(True)

    def _start_inventory_scan(self):
        src = self.inv_src_edit.text().strip()
        if not src or not os.path.isdir(src):
            QMessageBox.warning(self, "Invalid Directory", "Please select a valid existing directory to scan.")
            return

        out_base = self.inv_out_edit.text().strip()
        if not out_base:
            p = Path(src)
            fname = p.name or "File_Inventory"
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = Path(self.settings_store.value("default_output_dir", str(p.parent)) or p.parent)
            out_base = str(output_dir / f"{fname}_Inventory_{ts}")
            self.inv_out_edit.setText(out_base)

        if self.rb_both.isChecked():
            fmt = "both"
            excel_path = f"{out_base}.xlsx"
            html_path = f"{out_base}.html"
        elif self.rb_excel.isChecked():
            fmt = "excel"
            excel_path = f"{out_base}.xlsx"
            html_path = None
        else:
            fmt = "html"
            excel_path = None
            html_path = f"{out_base}.html"

        filter_cats = {
            "audio": self.cb_audio.isChecked(),
            "video": self.cb_video.isChecked(),
            "images": self.cb_images.isChecked(),
            "docs": self.cb_docs.isChecked(),
            "archives": self.cb_archives.isChecked(),
            "other": self.cb_other.isChecked(),
        }

        if not any(filter_cats.values()):
            QMessageBox.warning(self, "No Filters Selected", "Please select at least one file type category.")
            return

        enabled_cols = [cid for cid, cb in self.col_checkboxes.items() if cb.isChecked()]
        si_mode = "folder" if self.rb_si_folder.isChecked() else "continuous"

        config = {
            "source_folder": src,
            "output_excel": excel_path,
            "output_html": html_path,
            "format_type": fmt,
            "filter_categories": filter_cats,
            "enabled_columns": enabled_cols,
            "si_mode": si_mode,
            "skip_empty_folders": self.cb_skip_empty_folders.isChecked(),
            "size_unit": self.settings_store.value("inventory_size_unit", "auto"),
        }

        approved = self._approve_report_replacements([path for path in (excel_path, html_path) if path])
        if approved is None:
            return
        config['approved_replacements'] = approved
        config['protected_paths'] = self._report_evidence_targets()[0]

        # Clear previous UI results
        self._clear_results_actions()
        self.btn_run_inventory.setEnabled(False)
        self.btn_run_inventory.setText(" Processing Inventory")
        self.btn_run_inventory.setIcon(QIcon(self.assets["process"]))
        self.btn_run_inventory.setIconSize(QSize(16, 16))
        self.btn_cancel_inventory.setEnabled(True)
        if hasattr(self, "inv_spinner"):
            self.inv_spinner.start()
        self.inv_progress_bar.setRange(0, 0)  # Indeterminate at first
        self.inv_status_lbl.setText("Scanning folder tree for matching files...")

        self.start_time = time.time()
        self.timer_elapsed.start(100)

        # Launch QThread
        self.inventory_worker = InventoryWorker(config, self)
        self.inventory_worker.scan_progress.connect(self._on_inv_scan_progress)
        self.inventory_worker.process_started.connect(self._on_inv_process_started)
        self.inventory_worker.process_progress.connect(self._on_inv_process_progress)
        self.inventory_worker.writing_report.connect(self._on_inv_writing)
        self.inventory_worker.no_files_found.connect(self._on_inv_no_files)
        self.inventory_worker.error_occurred.connect(self._on_inv_error)
        self.inventory_worker.inventory_cancelled.connect(self._on_inv_cancelled)
        self.inventory_worker.inventory_finished.connect(self._on_inv_finished)
        self.inventory_worker.start()

    def _cancel_inventory_scan(self):
        if self.inventory_worker and self.inventory_worker.isRunning():
            self.inv_status_lbl.setText("Cancelling scan gracefully...")
            self.inventory_worker.cancel()

    def _update_elapsed_display(self):
        if self.start_time:
            el = time.time() - self.start_time
            self.stat_time.set_data(f"{el:.1f}s", "Active runtime")

    def _on_inv_scan_progress(self, total_files, folder_label):
        self.stat_files.set_data(f"{total_files:,}")
        self.inv_ticker_lbl.setText(f"Folder: {truncate_path(folder_label, 50)}")

    def _on_inv_process_started(self, total_files):
        self.inv_progress_bar.setRange(0, total_files)
        self.inv_progress_bar.setValue(0)
        self.inv_status_lbl.setText("Computing media durations & hashes...")

    def _on_inv_process_progress(self, current, total, fname):
        self.inv_progress_bar.setValue(current)
        self.inv_ticker_lbl.setText(f"Processing ({current}/{total}): {fname}")

    def _on_inv_writing(self, format_name):
        self.inv_progress_bar.setRange(0, 0)
        self.inv_status_lbl.setText(f"Formatting and writing {format_name}...")
        self.inv_ticker_lbl.setText("Applying styles, freezing panes, compiling metadata...")

    def _on_inv_no_files(self, errors):
        self._reset_inv_ui()
        self.inv_status_lbl.setText("Scan incomplete" if errors else "No matching files")
        if errors:
            QMessageBox.warning(self, "Scan Incomplete", "No inventory was generated.\n\n" + "\n".join(errors[:20]))
        else:
            QMessageBox.information(self, "No Matching Files", "No files matched your selected extension filters in this directory.")

    def _on_inv_error(self, title, msg):
        self._reset_inv_ui()
        QMessageBox.critical(self, title, msg)

    def _on_inv_cancelled(self, processed_count):
        self._reset_inv_ui()
        self.inv_status_lbl.setText(f"Cancelled by user after {processed_count:,} files.")
        QMessageBox.information(self, "Scan Cancelled", f"Operation was cancelled. Processed {processed_count:,} files.")

    def _on_inv_finished(self, generated_files, stats, errors):
        self._reset_inv_ui()
        self.last_generated_files = generated_files

        # Update KPI Cards
        self.stat_files.set_data(f"{stats.get('total_files', 0):,}", "Inventory items")
        self.stat_size.set_data(format_size(stats.get('total_size', 0)), "Total byte sum")
        self.stat_folders.set_data(f"{stats.get('total_folders', 0):,}", "Directories")
        self.stat_time.set_data(f"{stats.get('elapsed', 0.0):.2f}s", "Finished")

        self.inv_status_lbl.setText("Inventory completed with warnings" if errors else "Inventory Generation Complete!")
        self.inv_ticker_lbl.setText(f"Generated {len(generated_files)} report file(s) in {stats.get('elapsed', 0.0):.2f}s")
        self.inv_progress_bar.setRange(0, 100)
        self.inv_progress_bar.setValue(100)

        # Populate Result Cards
        self._populate_results_actions(generated_files)

        if errors:
            QMessageBox.warning(self, "Inventory Warnings", f"Reports generated with {len(errors)} warning(s). Some files or details may be missing.\n\n" + "\n".join(errors[:20]))
        else:
            dlg = CompletionDialog(generated_files, stats, parent=self)
            dlg.exec()

        # Log session activity for Overview telemetry
        src_path = self.inv_src_edit.text() if hasattr(self, "inv_src_edit") else "Directory"
        self._log_session_activity(
            "File Inventory",
            src_path,
            "Warnings" if errors else "Completed",
            f"Audited {stats.get('total_files', 0):,} files ({format_size(stats.get('total_size', 0))}) in {stats.get('elapsed', 0.0):.2f}s"
        )

    def _reset_inv_ui(self):
        self.timer_elapsed.stop()
        if hasattr(self, "inv_spinner"):
            self.inv_spinner.stop()
        self.btn_run_inventory.setEnabled(True)
        self.btn_run_inventory.setText("Run Inventory")
        self.btn_run_inventory.setIcon(QIcon())
        self.btn_cancel_inventory.setEnabled(False)
        self.inv_progress_bar.setRange(0, 100)

    def _clear_results_actions(self):
        while self.results_btn_layout.count():
            item = self.results_btn_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

    def _populate_results_actions(self, generated_files):
        self._clear_results_actions()
        if not generated_files:
            self.results_desc.setText("No output files generated.")
            return
        self.results_desc.setText("Reports ready.")
        for file_path in generated_files:
            path = Path(file_path).resolve()
            if path.suffix.lower() not in (".xlsx", ".html"):
                continue
            excel = path.suffix.lower() == ".xlsx"
            button = QPushButton("Open Excel" if excel else "Open HTML")
            button.setProperty("outputIcon", "file_spreadsheet" if excel else "file_code")
            button.setToolTip(str(path))
            button.setAccessibleDescription(str(path))
            button.clicked.connect(lambda checked=False, target=str(path): QDesktopServices.openUrl(QUrl.fromLocalFile(target)))
            self.results_btn_layout.addWidget(button)
        folder = str(Path(generated_files[0]).resolve().parent)
        button = QPushButton("Open Folder")
        button.setProperty("outputIcon", "folder")
        button.setToolTip(folder)
        button.setAccessibleDescription(folder)
        button.clicked.connect(lambda checked=False, target=folder: QDesktopServices.openUrl(QUrl.fromLocalFile(target)))
        self.results_btn_layout.addWidget(button)
        self._style_results_actions(self.current_theme)

    def _style_results_actions(self, theme):
        """Keep report actions consistent with the compact completion dialog."""
        light = theme == "light"
        background = "#f8fafc" if light else "#171e2e"
        foreground = "#172033" if light else "#edf2f9"
        border = "#cbd5e1" if light else "#344158"
        hover = "#eaf0f7" if light else "#243047"
        pressed = "#dde6f1" if light else "#2b3952"
        for index in range(self.results_btn_layout.count()):
            button = self.results_btn_layout.itemAt(index).widget()
            if not isinstance(button, QPushButton):
                continue
            icon = QIcon()
            for size in (16, 20, 24, 32, 40, 48):
                icon.addPixmap(LucideIcons.pixmap(button.property("outputIcon"), foreground, size))
            button.setIcon(icon)
            button.setIconSize(QSize(18, 18))
            button.setMinimumHeight(38)
            button.setCursor(QCursor(Qt.PointingHandCursor))
            button.setStyleSheet(f"""
                QPushButton {{ background: {background}; color: {foreground};
                    border: 1px solid {border}; border-radius: 5px;
                    padding: 8px 14px; font-family: 'Segoe UI';
                    font-size: 13px; font-weight: 500; }}
                QPushButton:hover {{ background: {hover}; border-color: #64748b; }}
                QPushButton:pressed {{ background: {pressed}; }}
                QPushButton:focus {{ border: 1px solid #60a5fa; }}
            """)

    # ==========================================================================
    # PAGE 2: CRYPTOGRAPHIC HASH VERIFICATION SUITE
    # ==========================================================================
    def _create_hash_page(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(14)

        # ── 1. Top Action Toolbar Card ──
        tb_card = SurfaceCard()
        tb_layout = QHBoxLayout()
        tb_layout.setSpacing(10)

        self.hash_btn_add_files = QPushButton("+ Add Files")
        self.hash_btn_add_files.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_primary(self.hash_btn_add_files, getattr(self, "current_theme", "dark"))
        self.hash_btn_add_files.clicked.connect(self._hash_add_files)
        tb_layout.addWidget(self.hash_btn_add_files)

        self.hash_btn_add_folder = QPushButton(" Add Folder")
        self.hash_btn_add_folder.setIcon(LucideIcons.icon("folder", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=16))
        self.hash_btn_add_folder.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_secondary(self.hash_btn_add_folder, getattr(self, "current_theme", "dark"))
        self.hash_btn_add_folder.clicked.connect(self._hash_add_folder)
        tb_layout.addWidget(self.hash_btn_add_folder)

        self.hash_btn_load_checksum = QPushButton(" Load Checksum")
        self.hash_btn_load_checksum.setIcon(LucideIcons.icon("file_text", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=16))
        self.hash_btn_load_checksum.setIconSize(QSize(16, 16))
        self.hash_btn_load_checksum.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_secondary(self.hash_btn_load_checksum, getattr(self, "current_theme", "dark"))
        self.hash_btn_load_checksum.clicked.connect(self._hash_load_checksum_manifest)
        tb_layout.addWidget(self.hash_btn_load_checksum)

        self.hash_btn_compare = QPushButton(" Compare 2 Files")
        self.hash_btn_compare.setIcon(LucideIcons.icon("compare", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=16))
        self.hash_btn_compare.setIconSize(QSize(16, 16))
        self.hash_btn_compare.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_secondary(self.hash_btn_compare, getattr(self, "current_theme", "dark"))
        self.hash_btn_compare.clicked.connect(self._open_compare_dialog)
        tb_layout.addWidget(self.hash_btn_compare)

        self.hash_btn_clear = QPushButton(" Clear")
        self.hash_btn_clear.setIcon(LucideIcons.icon("trash", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=16))
        self.hash_btn_clear.setIconSize(QSize(16, 16))
        self.hash_btn_clear.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_secondary(self.hash_btn_clear, getattr(self, "current_theme", "dark"))
        self.hash_btn_clear.clicked.connect(self._hash_clear_all)
        tb_layout.addWidget(self.hash_btn_clear)

        tb_layout.addStretch()

        tb_layout.addWidget(QLabel("New files:"))
        self.hash_algo_combo = QComboBox()
        self.hash_algo_combo.addItems(HASH_ALGORITHMS)
        self.hash_algo_combo.setCurrentText("SHA-256")
        self.hash_algo_combo.setToolTip("Default algorithm for newly added files. Existing rows retain the algorithm shown in their Algorithm column; applying a target detects its algorithm.")
        tb_layout.addWidget(self.hash_algo_combo)

        tb_card.main_layout.addLayout(tb_layout)

        # Row 2: Target Hash Input & Direct Setting Options
        target_row = QHBoxLayout()
        target_row.setSpacing(8)
        self.hash_target_lbl = QLabel("Target Hash:")
        self.hash_target_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #475569;" if getattr(self, "current_theme", "dark") == "light" else "font-size: 11px; font-weight: 700; color: #8e98b0;")
        target_lbl = self.hash_target_lbl
        self.target_hash_edit = QLineEdit()
        self.target_hash_edit.setPlaceholderText("Paste target/expected hash to verify against (or select row in table)...")
        self.target_hash_edit.returnPressed.connect(self._apply_target_hash)

        self.hash_btn_set_target = QPushButton("Set Target")
        self.hash_btn_set_target.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_primary(self.hash_btn_set_target, getattr(self, "current_theme", "dark"))
        self.hash_btn_set_target.clicked.connect(self._apply_target_hash)

        self.hash_btn_paste = QPushButton(" Paste")
        self.hash_btn_paste.setIcon(LucideIcons.icon("copy", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=14))
        self.hash_btn_paste.setIconSize(QSize(14, 14))
        self.hash_btn_paste.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_secondary(self.hash_btn_paste, getattr(self, "current_theme", "dark"))
        self.hash_btn_paste.clicked.connect(self._paste_and_apply_target)

        self.hash_btn_clear_tgt = QPushButton("Clear")
        self.hash_btn_clear_tgt.setCursor(QCursor(Qt.PointingHandCursor))
        self._style_button_secondary(self.hash_btn_clear_tgt, getattr(self, "current_theme", "dark"))
        self.hash_btn_clear_tgt.clicked.connect(self._clear_target_hash)

        target_row.addWidget(target_lbl)
        target_row.addWidget(self.target_hash_edit, 1)
        target_row.addWidget(self.hash_btn_set_target)
        target_row.addWidget(self.hash_btn_paste)
        target_row.addWidget(self.hash_btn_clear_tgt)
        tb_card.main_layout.addLayout(target_row)

        page_layout.addWidget(tb_card)

        # ── 2. Split Main View (Table + Hero/Inspector) ──
        split_layout = QHBoxLayout()
        split_layout.setSpacing(14)

        # LEFT: Hash Table with Search & Status Filter
        table_container = SurfaceCard(title="Files & Hashes", icon="file_text")
        
        # Search & Filter bar
        filter_row = QHBoxLayout()
        self.hash_search_edit = QLineEdit()
        self.hash_search_edit.setPlaceholderText("Search files or hashes...")
        self.hash_search_edit.textChanged.connect(self._filter_hash_table)
        filter_row.addWidget(self.hash_search_edit)

        self.hash_filter_combo = QComboBox()
        self.hash_filter_combo.addItem("All Items", "all")
        self.hash_filter_combo.addItem("Matches", "match")
        self.hash_filter_combo.addItem("Mismatches", "mismatch")
        self.hash_filter_combo.addItem("Calculated", "computed")
        self.hash_filter_combo.currentIndexChanged.connect(self._filter_hash_table)
        filter_row.addWidget(self.hash_filter_combo)
        table_container.main_layout.addLayout(filter_row)

        details_hint_row = QHBoxLayout()
        details_hint_row.setContentsMargins(0, 2, 0, 4)
        details_hint_row.setSpacing(6)
        self.hash_details_hint_icon = QLabel()
        self.hash_details_hint_icon.setFixedSize(16, 16)
        self.hash_details_hint = QLabel("Double-click the Algorithm cell in a calculated row to view all hash details.")
        self.hash_details_hint.setWordWrap(True)
        details_hint_row.addWidget(self.hash_details_hint_icon, 0, Qt.AlignTop)
        details_hint_row.addWidget(self.hash_details_hint, 1)
        table_container.main_layout.addLayout(details_hint_row)

        # Table Widget
        self.hash_table = QTableWidget()
        self.hash_table.setColumnCount(8)
        self.hash_table.setHorizontalHeaderLabels([
            "#", "Status", "File Name", "Size", "Algorithm", "Expected Hash", "Calculated Hash", "Result"
        ])
        self.hash_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.hash_table.horizontalHeader().setStretchLastSection(True)
        self.hash_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.hash_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.hash_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.hash_table.customContextMenuRequested.connect(self._on_hash_table_context_menu)
        self.hash_table.itemDoubleClicked.connect(self._on_hash_table_double_click)
        self.hash_table.itemSelectionChanged.connect(self._on_hash_table_selection_changed)
        table_container.main_layout.addWidget(self.hash_table)

        table_container.setMinimumWidth(360)
        split_layout.addWidget(table_container, 58)

        # RIGHT: Hero Animation & Controls Card (Flexible Scrollable Container)
        right_hash_scroll = QScrollArea()
        right_hash_scroll.setWidgetResizable(True)
        right_hash_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        right_hash_scroll.setFrameShape(QFrame.NoFrame)
        right_hash_scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        right_panel = SurfaceCard(title="Verification", icon="shield")
        right_panel.setMinimumWidth(280)

        # Animated Verification Radar / Shield Widget
        self.verification_loader = AnimatedVerificationLoader()
        right_panel.main_layout.addWidget(self.verification_loader)

        # KPI Metrics Row
        kpi_row = QGridLayout()
        kpi_row.setSpacing(8)
        self.kpi_total = ModernStatCard("file_text", "Files", "0")
        self.kpi_matches = ModernStatCard("check_circle", "Matches", "0")
        self.kpi_mismatches = ModernStatCard("x_circle", "Mismatches", "0")
        self.kpi_unchecked = ModernStatCard("zap", "Calculated", "0")

        kpi_row.addWidget(self.kpi_total, 0, 0)
        kpi_row.addWidget(self.kpi_matches, 0, 1)
        kpi_row.addWidget(self.kpi_mismatches, 1, 0)
        kpi_row.addWidget(self.kpi_unchecked, 1, 1)
        right_panel.main_layout.addLayout(kpi_row)

        # Big Action Verification Buttons
        # Action Buttons row: Verify Checksums + Cancel side-by-side with appropriate short widths
        verify_btn_row = QHBoxLayout()
        verify_btn_row.setSpacing(8)

        self.btn_verify_hashes = QPushButton("Verify Checksums")
        self.btn_verify_hashes.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_verify_hashes.setFixedWidth(170)
        self.btn_verify_hashes.setFixedHeight(36)
        self.btn_verify_hashes.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: #ffffff;
                font-size: 12px;
                font-weight: 700;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #60a5fa; }
            QPushButton:pressed { background-color: #2563eb; }
        """)
        self.btn_verify_hashes.clicked.connect(self._start_hash_verification)
        verify_btn_row.addStretch()
        verify_btn_row.addWidget(self.btn_verify_hashes)

        self.btn_cancel_verify = QPushButton("Cancel")
        self.btn_cancel_verify.setEnabled(False)
        self.btn_cancel_verify.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_cancel_verify.setFixedWidth(85)
        self.btn_cancel_verify.setFixedHeight(36)
        self.btn_cancel_verify.setStyleSheet("""
            QPushButton {
                background-color: #ef4444;
                color: #ffffff;
                font-size: 12px;
                font-weight: 700;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #f87171; }
            QPushButton:disabled { background-color: #1e243b; color: #55607c; }
        """)
        self.btn_cancel_verify.clicked.connect(self._cancel_hash_verification)
        verify_btn_row.addWidget(self.btn_cancel_verify)
        verify_btn_row.addStretch()
        right_panel.main_layout.addLayout(verify_btn_row)

        # Quick Compare 2 Files inside Right Panel (Standard size, centered)
        self.hash_btn_compare2 = QPushButton(" Compare 2 Files")
        self.hash_btn_compare2.setIcon(LucideIcons.icon("compare", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=16))
        self.hash_btn_compare2.setIconSize(QSize(16, 16))
        self.hash_btn_compare2.setCursor(QCursor(Qt.PointingHandCursor))
        self.hash_btn_compare2.setFixedWidth(180)
        self.hash_btn_compare2.setFixedHeight(34)
        self._style_button_secondary(self.hash_btn_compare2, getattr(self, "current_theme", "dark"))
        self.hash_btn_compare2.clicked.connect(self._open_compare_dialog)
        cmp2_row = QHBoxLayout()
        cmp2_row.addStretch()
        cmp2_row.addWidget(self.hash_btn_compare2)
        cmp2_row.addStretch()
        right_panel.main_layout.addLayout(cmp2_row)

        # Export Actions Dropdown / Buttons
        self.hash_export_lbl = QLabel("EXPORT:")
        self.hash_export_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #475569; margin-top: 4px;" if getattr(self, "current_theme", "dark") == "light" else "font-size: 11px; font-weight: 700; color: #8e98b0; margin-top: 4px;")
        right_panel.main_layout.addWidget(self.hash_export_lbl)

        export_grid = QGridLayout()
        export_grid.setSpacing(6)

        exp_manifest_btn = QPushButton(" Checksum")
        exp_manifest_btn.setIcon(LucideIcons.icon("file_text", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=14))
        exp_manifest_btn.setIconSize(QSize(14, 14))
        exp_manifest_btn.clicked.connect(self._export_hash_manifest)

        exp_excel_btn = QPushButton(" Excel")
        exp_excel_btn.setIcon(LucideIcons.icon("file_spreadsheet", color="#287a45" if getattr(self, "current_theme", "dark") == "light" else ("#80cc93").replace("#80cc93", status_green(self)), size=14))
        exp_excel_btn.setIconSize(QSize(14, 14))
        exp_excel_btn.clicked.connect(self._export_hash_excel)
        self.exp_excel_btn = exp_excel_btn

        exp_html_btn = QPushButton(" HTML")
        exp_html_btn.setIcon(LucideIcons.icon("file_code", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=14))
        exp_html_btn.setIconSize(QSize(14, 14))
        exp_html_btn.clicked.connect(self._export_hash_html)

        exp_csv_btn = QPushButton(" CSV")
        exp_csv_btn.setIcon(LucideIcons.icon("database", color="#2563eb" if getattr(self, "current_theme", "dark") == "light" else "#60a5fa", size=14))
        exp_csv_btn.setIconSize(QSize(14, 14))
        exp_csv_btn.clicked.connect(self._export_hash_csv)

        self.hash_export_buttons = [exp_manifest_btn, exp_excel_btn, exp_html_btn, exp_csv_btn]
        for b in self.hash_export_buttons:
            b.setCursor(QCursor(Qt.PointingHandCursor))
            self._style_button_secondary(b, getattr(self, "current_theme", "dark"))

        export_grid.addWidget(exp_manifest_btn, 0, 0)
        export_grid.addWidget(exp_excel_btn, 0, 1)
        export_grid.addWidget(exp_html_btn, 1, 0)
        export_grid.addWidget(exp_csv_btn, 1, 1)
        right_panel.main_layout.addLayout(export_grid)

        right_panel.main_layout.addStretch()
        right_hash_scroll.setWidget(right_panel)
        split_layout.addWidget(right_hash_scroll, 42)

        page_layout.addLayout(split_layout)

        return page

    # ==========================================================================
    # HASH VERIFICATION ACTIONS & SLOTS
    # ==========================================================================
    def _hash_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files for Cryptographic Hashing")
        if files:
            for f in files:
                self._add_file_to_hash_items(f)
            self._refresh_hash_table()

    def _hash_add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Hash Recursively")
        if folder:
            scan_errors = []
            for root, _, files in evidence_walk(folder, onerror=lambda error: scan_errors.append(str(error))):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    if os.path.isfile(fpath):
                        self._add_file_to_hash_items(fpath)
            self._refresh_hash_table()
            if scan_errors:
                QMessageBox.warning(self, "Folder Scan", "Some paths were skipped:\n" + "\n".join(scan_errors[:20]))

    def _add_file_to_hash_items(self, file_path, expected_hash="", algorithm=None):
        p = Path(file_path)
        try:
            size = p.stat().st_size
        except Exception:
            size = 0

        item = {
            "id": str(uuid.uuid4()),
            "name": p.name,
            "path": str(p.resolve()),
            "size": size,
            "algorithm": algorithm or self.hash_algo_combo.currentText(),
            "expected_hash": expected_hash,
            "calculated_hash": "",
            "all_hashes": {},
            "status": "pending",
            "status_text": "Pending",
            "result": "PENDING" if not expected_hash else "UNCHECKED",
            "error": None
        }
        self.hash_items.append(item)

    def _hash_load_checksum_manifest(self):
        chk_file, _ = QFileDialog.getOpenFileName(
            self, "Load Checksum Manifest File", "",
            "Checksum Files (*.sha256 *.md5 *.sha1 *.sha512 *.txt);;All Files (*.*)"
        )
        if not chk_file:
            return

        base_dir = Path(chk_file).parent
        loaded_count = 0
        unresolved = []

        try:
            with open(chk_file, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    parts = line.split(None, 1)
                    if len(parts) == 2:
                        raw_hash = parts[0].strip()
                        raw_name = parts[1].strip().lstrip("*")
                        cand_path = base_dir / raw_name
                        detected_algo = detect_hash_algorithm(raw_hash)
                        if cand_path.is_file() and detected_algo:
                            self._add_file_to_hash_items(
                                str(cand_path),
                                expected_hash=raw_hash,
                                algorithm=detected_algo or self.hash_algo_combo.currentText()
                            )
                            loaded_count += 1
                        else:
                            unresolved.append(raw_name)
            self._refresh_hash_table()
            QMessageBox.information(
                self, "Checksums Loaded",
                f"Imported {loaded_count} files. {len(unresolved)} missing or invalid entries." + ("\n\n" + "\n".join(unresolved[:20]) if unresolved else "")
            )
        except Exception as ex:
            QMessageBox.critical(self, "Load Error", f"Failed to read checksum file:\n{ex}")

    def _hash_clear_all(self):
        if not self.hash_items:
            return
        self.hash_items.clear()
        self._refresh_hash_table()
        self.verification_loader.set_state("idle", "Integrity Engine Ready", "Manifest cleared")
        self.kpi_total.set_data("0")
        self.kpi_matches.set_data("0")
        self.kpi_mismatches.set_data("0")
        self.kpi_unchecked.set_data("0")

    def _refresh_hash_table(self):
        query = self.hash_search_edit.text().strip().lower()
        filter_mode = self.hash_filter_combo.currentData()

        # Update KPI cards
        total = len(self.hash_items)
        matches = sum(1 for it in self.hash_items if it.get("result") == "MATCH")
        mismatches = sum(1 for it in self.hash_items if it.get("result") == "MISMATCH")
        computed = sum(1 for it in self.hash_items if it.get("result") == "NO TARGET")

        self.kpi_total.set_data(f"{total:,}")
        self.kpi_matches.set_data(f"{matches:,}")
        self.kpi_mismatches.set_data(f"{mismatches:,}")
        self.kpi_unchecked.set_data(f"{computed:,}")

        visible_items = []
        for it in self.hash_items:
            if filter_mode == "match" and it.get("result") != "MATCH":
                continue
            if filter_mode == "mismatch" and it.get("result") != "MISMATCH":
                continue
            if filter_mode == "computed" and it.get("result") != "NO TARGET":
                continue

            if query:
                combined_txt = f"{it['name']} {it.get('expected_hash', '')} {it.get('calculated_hash', '')} {it.get('status_text', '')}".lower()
                if query not in combined_txt:
                    continue

            visible_items.append(it)

        self._hash_result_counts = {'MATCH': matches, 'MISMATCH': mismatches, 'NO TARGET': computed}
        self._hash_row_by_id = {}
        self.hash_table.setRowCount(len(visible_items))
        for row_idx, item in enumerate(visible_items):
            self._hash_row_by_id[item['id']] = row_idx
            self._populate_hash_row(row_idx, item)

    def _populate_hash_row(self, row_idx, item):
        # Column 0: SI No
        si_item = QTableWidgetItem(str(row_idx + 1))
        si_item.setTextAlignment(Qt.AlignCenter)
        self.hash_table.setItem(row_idx, 0, si_item)

        # Column 1: Status Badge
        status_text = item.get("status_text", "Pending")
        st_item = QTableWidgetItem(status_text)
        st_item.setTextAlignment(Qt.AlignCenter)
        res = item.get("result", "")
        if res == "MATCH":
            st_item.setForeground(QColor(("#80cc93").replace("#80cc93", status_green(self))))
        elif res == "MISMATCH":
            st_item.setForeground(QColor("#f87171"))
        elif res == "NO TARGET":
            st_item.setForeground(QColor("#60a5fa"))
        self.hash_table.setItem(row_idx, 1, st_item)

        # Column 2: File Name
        name_item = QTableWidgetItem(item["name"])
        name_item.setToolTip(item["path"])
        name_item.setData(Qt.UserRole, item["id"])
        self.hash_table.setItem(row_idx, 2, name_item)

        # Column 3: File Size
        size_item = QTableWidgetItem(format_size(item.get("size", 0)))
        size_item.setTextAlignment(Qt.AlignCenter)
        self.hash_table.setItem(row_idx, 3, size_item)

        # Column 4: Algorithm
        algo_item = QTableWidgetItem(item.get("algorithm", "SHA-256"))
        algo_item.setTextAlignment(Qt.AlignCenter)
        self.hash_table.setItem(row_idx, 4, algo_item)

        # Column 5: Expected Hash
        exp_item = QTableWidgetItem(item.get("expected_hash") or "—")
        exp_item.setFont(QFont("Consolas", 9))
        self.hash_table.setItem(row_idx, 5, exp_item)

        # Column 6: Calculated Hash
        calc_item = QTableWidgetItem(item.get("calculated_hash") or "—")
        calc_item.setFont(QFont("Consolas", 9))
        self.hash_table.setItem(row_idx, 6, calc_item)

        # Column 7: Result
        res_item = QTableWidgetItem(item.get("result", "PENDING"))
        res_item.setTextAlignment(Qt.AlignCenter)
        if res == "MATCH":
            res_item.setForeground(QColor(("#80cc93").replace("#80cc93", status_green(self))))
        elif res == "MISMATCH":
            res_item.setForeground(QColor("#f87171"))
        elif res == "NO TARGET":
            res_item.setForeground(QColor("#60a5fa"))
        self.hash_table.setItem(row_idx, 7, res_item)

    def _filter_hash_table(self):
        self._refresh_hash_table()

    def _detect_hash_algorithm(self, hash_str: str):
        if not hash_str:
            return None
        h = hash_str.strip().lower()
        if len(h) == 32 and all(c in "0123456789abcdef" for c in h):
            return "MD5"
        if len(h) == 40 and all(c in "0123456789abcdef" for c in h):
            return "SHA-1"
        if len(h) == 64 and all(c in "0123456789abcdef" for c in h):
            return "SHA-256"
        if len(h) == 128 and all(c in "0123456789abcdef" for c in h):
            return "SHA-512"
        return None

    def _reevaluate_item_match(self, item: dict):
        expected = (item.get("expected_hash") or "").strip()
        algo = item.get("algorithm") or self.hash_algo_combo.currentText()
        all_h = item.get("all_hashes")
        if all_h:
            item["calculated_hash"] = all_h.get(algo, "")
        elif item.get("algorithm") != detect_hash_algorithm(item.get("calculated_hash") or ""):
            item["calculated_hash"] = ""

        calc = (item.get("calculated_hash") or "").strip()
        if calc and calc != "—":
            if expected:
                if calc.lower() == expected.lower():
                    item["status"] = "match"
                    item["status_text"] = "Verified Match"
                    item["result"] = "MATCH"
                else:
                    item["status"] = "mismatch"
                    item["status_text"] = "Mismatch!"
                    item["result"] = "MISMATCH"
            else:
                item["status"] = "computed"
                item["status_text"] = "Calculated"
                item["result"] = "NO TARGET"
        else:
            if expected:
                item["status"] = "pending"
                item["status_text"] = "Target Set"
                item["result"] = "PENDING"
            else:
                item["status"] = "pending"
                item["status_text"] = "Pending"
                item["result"] = "PENDING"

    def _on_hash_table_selection_changed(self):
        selected_indexes = self.hash_table.selectedIndexes()
        if not selected_indexes:
            return
        row = selected_indexes[0].row()
        name_cell = self.hash_table.item(row, 2)
        if not name_cell:
            return
        item_id = name_cell.data(Qt.UserRole)
        target = next((it for it in self.hash_items if it["id"] == item_id), None)
        if target:
            exp = target.get("expected_hash") or ""
            if not self.target_hash_edit.hasFocus():
                self.target_hash_edit.setText(exp)

    def _apply_target_hash(self):
        val = self.target_hash_edit.text().strip()
        if not val:
            QMessageBox.information(self, "Target Hash", "Please enter or paste a target checksum first.")
            return

        if not self.hash_items:
            QMessageBox.information(self, "No Files", "Please add files to verify before setting a target hash.")
            return

        selected_rows = sorted(list(set(idx.row() for idx in self.hash_table.selectedIndexes())))
        target_items = []
        if selected_rows:
            for r in selected_rows:
                name_cell = self.hash_table.item(r, 2)
                if name_cell:
                    item_id = name_cell.data(Qt.UserRole)
                    item = next((it for it in self.hash_items if it["id"] == item_id), None)
                    if item:
                        target_items.append(item)
        else:
            target_items = [self.hash_items[0]]

        detected_algo = self._detect_hash_algorithm(val)
        if detected_algo:
            self.hash_algo_combo.setCurrentText(detected_algo)

        for it in target_items:
            it["expected_hash"] = val.lower()
            if detected_algo:
                it["algorithm"] = detected_algo
            self._reevaluate_item_match(it)

        self._refresh_hash_table()

    def _paste_and_apply_target(self):
        clip_text = QApplication.clipboard().text().strip()
        if not clip_text:
            QMessageBox.information(self, "Clipboard Empty", "Your clipboard does not contain any text.")
            return
        if ":" in clip_text and len(clip_text.split(":", 1)[1].strip()) in (32, 40, 64, 128):
            clip_text = clip_text.split(":", 1)[1].strip()
        self.target_hash_edit.setText(clip_text)
        self._apply_target_hash()

    def _clear_target_hash(self):
        self.target_hash_edit.clear()
        selected_rows = sorted(list(set(idx.row() for idx in self.hash_table.selectedIndexes())))
        target_items = []
        if selected_rows:
            for r in selected_rows:
                name_cell = self.hash_table.item(r, 2)
                if name_cell:
                    item_id = name_cell.data(Qt.UserRole)
                    item = next((it for it in self.hash_items if it["id"] == item_id), None)
                    if item:
                        target_items.append(item)
        else:
            target_items = self.hash_items

        for it in target_items:
            it["expected_hash"] = ""
            self._reevaluate_item_match(it)

        self._refresh_hash_table()

    def _prompt_set_target_hash_for_row(self, row: int):
        name_cell = self.hash_table.item(row, 2)
        if not name_cell:
            return
        item_id = name_cell.data(Qt.UserRole)
        item = next((it for it in self.hash_items if it["id"] == item_id), None)
        if not item:
            return

        curr_target = item.get("expected_hash") or ""
        new_val, ok = QInputDialog.getText(
            self,
            "Set Target Hash",
            f"Enter expected checksum for:\n{item['name']}",
            text=curr_target
        )
        if ok:
            val = new_val.strip()
            item["expected_hash"] = val.lower() if val else ""
            detected_algo = self._detect_hash_algorithm(val)
            if detected_algo:
                item["algorithm"] = detected_algo
            self._reevaluate_item_match(item)
            self._refresh_hash_table()

    def _on_hash_table_double_click(self, item):
        row = item.row()
        name_cell = self.hash_table.item(row, 2)
        if not name_cell:
            return
        item_id = name_cell.data(Qt.UserRole)
        target = next((it for it in self.hash_items if it["id"] == item_id), None)
        if target:
            if item.column() == 5:
                self._prompt_set_target_hash_for_row(row)
            else:
                dlg = HashInspectorDialog(target, parent=self)
                dlg.exec()

    def _on_hash_table_context_menu(self, pos):
        item_at = self.hash_table.itemAt(pos)
        if not item_at:
            return
        row = item_at.row()
        name_cell = self.hash_table.item(row, 2)
        if not name_cell:
            return
        item_id = name_cell.data(Qt.UserRole)
        target = next((it for it in self.hash_items if it["id"] == item_id), None)
        if not target:
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #161b2b;
                border: 1px solid #28334e;
                color: #f1f5f9;
                padding: 6px;
            }
            QMenu::item {
                padding: 6px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #2563eb;
                color: #ffffff;
            }
        """)

        act_set_target = menu.addAction(QIcon(self.assets["filter"]), "Set Target Hash...")
        act_paste_target = menu.addAction(QIcon(self.assets["paste"]), "Paste Target from Clipboard")
        act_clear_target = menu.addAction(QIcon(self.assets["trash"]), "Clear Target Hash")
        menu.addSeparator()
        act_inspect = menu.addAction(QIcon(self.assets["inspect"]), "Inspect All Hashes (MD5, SHA-1, SHA-256, SHA-512)...")
        act_copy_calc = menu.addAction(QIcon(self.assets["paste"]), "Copy Calculated Hash")
        menu.addSeparator()
        act_compare = menu.addAction(QIcon(self.assets["compare"]), "Compare with Another File...")
        act_remove = menu.addAction(QIcon(self.assets["cross_badge"]), "Remove from List")

        chosen = menu.exec(self.hash_table.viewport().mapToGlobal(pos))
        if chosen == act_set_target:
            self._prompt_set_target_hash_for_row(row)
        elif chosen == act_paste_target:
            clip = QApplication.clipboard().text().strip()
            if clip:
                if ":" in clip and len(clip.split(":", 1)[1].strip()) in (32, 40, 64, 128):
                    clip = clip.split(":", 1)[1].strip()
                target["expected_hash"] = clip.lower()
                detected = self._detect_hash_algorithm(clip)
                if detected:
                    target["algorithm"] = detected
                self._reevaluate_item_match(target)
                self._refresh_hash_table()
        elif chosen == act_clear_target:
            target["expected_hash"] = ""
            self._reevaluate_item_match(target)
            self._refresh_hash_table()
        elif chosen == act_inspect:
            dlg = HashInspectorDialog(target, parent=self)
            dlg.exec()
        elif chosen == act_copy_calc:
            calc_h = target.get("calculated_hash") or ""
            if calc_h:
                QApplication.clipboard().setText(calc_h)
                QMessageBox.information(self, "Copied", "Calculated hash copied to clipboard!")
        elif chosen == act_compare:
            dlg = CompareTwoFilesDialog(file1=target.get("path"), parent=self)
            dlg.exec()
        elif chosen == act_remove:
            self.hash_items = [it for it in self.hash_items if it["id"] != item_id]
            self._refresh_hash_table()

    def _start_hash_verification(self):
        if not self.hash_items:
            QMessageBox.information(self, "No Items", "Please add files or folders to verify first.")
            return

        default_algo = self.hash_algo_combo.currentText()
        self.btn_verify_hashes.setEnabled(False)
        self.btn_verify_hashes.setText(" Processing Checksums")
        self.btn_verify_hashes.setIcon(QIcon(self.assets["process"]))
        self.btn_verify_hashes.setIconSize(QSize(16, 16))
        self.btn_cancel_verify.setEnabled(True)

        self._set_hash_controls_enabled(False)
        for item in self.hash_items:
            item.update(calculated_hash="", all_hashes={}, status="pending", result="PENDING", error=None)
        self._refresh_hash_table()
        self.hash_worker = HashVerificationWorker(self.hash_items, default_algo, self)
        self.hash_worker.verification_cancelled.connect(self._on_hash_cancelled)
        self.hash_worker.verification_completed.connect(self._on_hash_completed)
        self._hash_models_by_id = {item['id']: item for item in self.hash_items}
        self._pending_hash_terminal = None
        self._last_hash_progress = None
        if not hasattr(self, '_hash_update_timer'):
            self._hash_update_timer = QTimer(self)
            self._hash_update_timer.setTimerType(Qt.TimerType.PreciseTimer)
            self._hash_update_timer.timeout.connect(self._drain_hash_updates)
        self._hash_update_timer.start(33)
        self.verification_loader.set_state('verifying', 'Computing Checksums...', 'Reading selected files')
        self.hash_worker.start()

    def _cancel_hash_verification(self):
        if self.hash_worker and self.hash_worker.isRunning():
            self.verification_loader.set_state("idle", "Cancelling Verification...", "Halting hash threads")
            self.hash_worker.cancel()

    def _on_hash_item_started(self, file_id, idx, total, name):
        self.verification_loader.set_state(
            "verifying",
            f"Hashing ({idx}/{total})",
            truncate_path(name, 35)
        )

    def _on_hash_item_completed(self, file_id, updated_item):
        model = getattr(self, '_hash_models_by_id', {}).get(file_id)
        if model is None:
            self._hash_models_by_id = {item['id']: item for item in self.hash_items}
            model = self._hash_models_by_id.get(file_id)
        if model is None:
            return
        counts = self._hash_result_counts
        previous = model.get('result')
        if previous in counts:
            counts[previous] -= 1
        model.update(updated_item)
        result = model.get('result')
        if result in counts:
            counts[result] += 1
        query = self.hash_search_edit.text().strip().lower()
        mode = self.hash_filter_combo.currentData()
        visible = not ((mode == 'match' and result != 'MATCH') or
                       (mode == 'mismatch' and result != 'MISMATCH') or
                       (mode == 'computed' and result != 'NO TARGET'))
        text = f"{model['name']} {model.get('expected_hash', '')} {model.get('calculated_hash', '')} {model.get('status_text', '')}".lower()
        visible = visible and (not query or query in text)
        row = self._hash_row_by_id.get(file_id)
        if row is None and visible:
            row = self.hash_table.rowCount()
            self.hash_table.insertRow(row)
            self._hash_row_by_id[file_id] = row
        if row is not None:
            self.hash_table.setRowHidden(row, not visible)
            if visible:
                self._populate_hash_row(row, model)

    def _drain_hash_updates(self):
        worker = self.hash_worker
        if worker is not None:
            progress = worker.ui_progress
            if progress and progress != getattr(self, '_last_hash_progress', None):
                self._on_hash_item_started(*progress)
                self._last_hash_progress = progress
            deadline = time.perf_counter() + .004
            for _ in range(32):
                if time.perf_counter() >= deadline:
                    break
                try:
                    file_id, item = worker.ui_results.get_nowait()
                except queue.Empty:
                    break
                self._on_hash_item_completed(file_id, item)
            counts = self._hash_result_counts
            for widget, result in ((self.kpi_matches, 'MATCH'), (self.kpi_mismatches, 'MISMATCH'), (self.kpi_unchecked, 'NO TARGET')):
                widget.set_data(f"{counts[result]:,}")
        terminal = getattr(self, '_pending_hash_terminal', None)
        if terminal and (worker is None or worker.ui_results.empty()):
            if hasattr(self, '_hash_update_timer'):
                self._hash_update_timer.stop()
            self._pending_hash_terminal = None
            mode, result = terminal
            if mode == 'completed':
                self._finish_hash_completed(result)
            else:
                self._finish_hash_cancelled(result)


    def _on_hash_cancelled(self, processed_count):
        self._pending_hash_terminal = ('cancelled', processed_count)
        self._drain_hash_updates()

    def _finish_hash_cancelled(self, processed_count):
        self._set_hash_controls_enabled(True)
        self.btn_verify_hashes.setEnabled(True)
        self.btn_verify_hashes.setText("Verify Checksums")
        self.btn_verify_hashes.setIcon(QIcon())
        self.btn_cancel_verify.setEnabled(False)
        self.verification_loader.set_state("idle", "Verification Stopped", f"Processed {processed_count} items")
        QMessageBox.information(self, "Stopped", f"Verification stopped by user after {processed_count} files.")

    def _on_hash_completed(self, summary):
        self._pending_hash_terminal = ('completed', summary)
        self._drain_hash_updates()

    def _finish_hash_completed(self, summary):
        self._set_hash_controls_enabled(True)
        self.btn_verify_hashes.setEnabled(True)
        self.btn_verify_hashes.setText("Verify Checksums")
        self.btn_verify_hashes.setIcon(QIcon())
        self.btn_cancel_verify.setEnabled(False)
        matched = summary.get("matched", 0)
        errors = summary.get("errors", 0)
        mismatched = summary.get("mismatched", 0)
        total = summary.get("total", 0)
        details = f"{matched} matched, {mismatched} mismatched, {errors} errors, {summary.get('computed', 0)} calculated without targets"
        if errors or mismatched:
            state, title, status = "mismatch", "Verification incomplete or failed", "Alert"
        elif total and matched == total:
            state, title, status = "match", "ALL EXPECTED HASHES MATCH", "Verified"
        else:
            state, title, status = "idle", "Calculation Finished", "Completed"
        self.verification_loader.set_state(state, title, details)
        self._log_session_activity("Hash Verification", f"{total} items", status, details)

    @report_export_errors
    def _export_hash_manifest(self):
        if not self.hash_items:
            QMessageBox.information(self, "No Items", "No hash items to export.")
            return
        f, _ = QFileDialog.getSaveFileName(self, "Save Checksum File", self._default_report_file("checksums.sha256"), "Checksum Files (*.sha256 *.md5)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            HashReportExporter.export_checksum_file(f, self.hash_items, self.hash_algo_combo.currentText())
            QMessageBox.information(self, "Exported", f"Saved checksum file:\n{f}")

    @report_export_errors
    def _export_hash_excel(self):
        if not self.hash_items:
            QMessageBox.information(self, "No Items", "No hash items to export.")
            return
        f, _ = QFileDialog.getSaveFileName(self, "Save Verification Excel Report", self._default_report_file("Hash_Verification.xlsx"), "Excel Files (*.xlsx)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            HashReportExporter.export_excel(f, self.hash_items, self.hash_algo_combo.currentText())
            QMessageBox.information(self, "Exported", f"Saved Excel report:\n{f}")

    @report_export_errors
    def _export_hash_html(self):
        if not self.hash_items:
            QMessageBox.information(self, "No Items", "No hash items to export.")
            return
        f, _ = QFileDialog.getSaveFileName(self, "Save Verification HTML Report", self._default_report_file("Hash_Verification.html"), "HTML Files (*.html)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            HashReportExporter.export_html(f, self.hash_items, self.hash_algo_combo.currentText())
            QMessageBox.information(self, "Exported", f"Saved HTML report:\n{f}")

    @report_export_errors
    def _export_hash_csv(self):
        if not self.hash_items:
            QMessageBox.information(self, "No Items", "No hash items to export.")
            return
        f, _ = QFileDialog.getSaveFileName(self, "Save Verification CSV", self._default_report_file("Hash_Verification.csv"), "CSV Files (*.csv)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            HashReportExporter.export_csv(f, self.hash_items, self.hash_algo_combo.currentText())
            QMessageBox.information(self, "Exported", f"Saved CSV report:\n{f}")

    # ==========================================================================
    # SECTION 7.3: METADATA EXPLORER (EXIFTOOL INTEGRATION) UI & CONTROLLER
    # ==========================================================================
    def _create_metadata_page(self):
        """
        Constructs the third main module: Metadata Explorer.
        Features a spacious, minimal layout:
        - Left Panel:
          1. File & Folder Selection card (Single File / Folder, path input, browse buttons,
             drag & drop zone, SmallGalleryView file preview, detail info, include subfolders checkbox).
          2. Generated Reports card (under the gallery view with quick open buttons for HTML, Text, JSON,
             and a prominent 'Open Output Folder' button).
          3. Action buttons (Extract Metadata & Cancel) horizontally centered.
        - Right Panel:
          1. Top Container: Extraction Settings card (Hashes, integrity, raw JSON, report formats, output folder, ExifTool ready status).
          2. Main Container: Metadata Viewer card expanded all the way down for maximum visibility
             (live search, category filter, full-height metadata tree, Expand/Collapse/Copy/View Raw JSON).
        """
        page = QWidget()
        page_layout = QHBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        # ──────────────────────────────────────────────────────────────────────
        # LEFT PANEL: SELECTION, GALLERY PREVIEW, REPORTS & ACTIONS
        # ──────────────────────────────────────────────────────────────────────
        left_col_widget = QWidget()
        left_col_widget.setMinimumWidth(380)
        left_col_layout = QVBoxLayout(left_col_widget)
        left_col_layout.setContentsMargins(0, 0, 8, 0)
        left_col_layout.setSpacing(8)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setStyleSheet("QScrollArea { background-color: transparent; border: none; }")

        left_container = QWidget()
        left_container.setObjectName("LeftContainer")
        left_container.setStyleSheet("#LeftContainer { background-color: transparent; }")
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 6, 0)
        left_layout.setSpacing(8)

        # 1. File & Folder Selection Card
        ev_card = SurfaceCard(title="Select File or Folder", icon="folder")

        # Mode Selection: Single File vs Folder
        mode_row = QHBoxLayout()
        mode_lbl = QLabel("Target Type:")
        mode_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #8e98b0;")
        mode_row.addWidget(mode_lbl)

        self.meta_rb_file = QRadioButton("Single File")
        self.meta_rb_folder = QRadioButton("Folder")
        self.meta_rb_file.setChecked(True)
        self.meta_target_group = QButtonGroup(self)
        self.meta_target_group.addButton(self.meta_rb_file)
        self.meta_target_group.addButton(self.meta_rb_folder)
        self.meta_rb_file.toggled.connect(self._on_meta_target_mode_changed)
        mode_row.addWidget(self.meta_rb_file)
        mode_row.addWidget(self.meta_rb_folder)
        mode_row.addStretch()
        ev_card.main_layout.addLayout(mode_row)

        # Path input and responsive browse button
        path_row = QHBoxLayout()
        path_row.setSpacing(6)
        self.meta_src_edit = QLineEdit()
        self.meta_src_edit.setPlaceholderText("Select a file or folder...")
        self.meta_src_edit.textChanged.connect(self._update_meta_evidence_display)

        self.meta_btn_browse = QPushButton("Browse")
        self.meta_btn_browse.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_browse.setStyleSheet("""
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 8px;
                color: #f1f5f9;
                padding: 8px 14px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
        """)
        self.meta_btn_browse.clicked.connect(self._on_meta_browse_clicked)
        self.meta_btn_browse_file = self.meta_btn_browse
        self.meta_btn_browse_folder = self.meta_btn_browse

        path_row.addWidget(self.meta_src_edit)
        path_row.addWidget(self.meta_btn_browse)
        ev_card.main_layout.addLayout(path_row)

        # Drag and drop target zone
        self.meta_drop_zone = FileDropWidget(parent=self)
        self.meta_drop_zone.path_dropped.connect(self._on_meta_path_dropped)
        ev_card.main_layout.addWidget(self.meta_drop_zone)

        # Folder subdirectories recursive checkbox
        self.meta_cb_recursive = QCheckBox("Include subfolders")
        self.meta_cb_recursive.setChecked(False)
        self.meta_cb_recursive.setEnabled(False)
        self.meta_cb_recursive.toggled.connect(lambda: self._update_meta_evidence_display())
        ev_card.main_layout.addWidget(self.meta_cb_recursive)

        # Hidden evidence detail info label for backward compatibility
        self.meta_info_lbl = QLabel("")
        self.meta_info_lbl.hide()

        # 2. Dynamic Preview / Gallery Viewer (Automatically expands to fill vertical space)
        self.meta_preview_gallery = DynamicPreviewGalleryViewer(parent=self)
        self.meta_gallery_view = self.meta_preview_gallery
        self.meta_preview_widget = self.meta_preview_gallery
        self.meta_preview_gallery.file_selected.connect(self._on_gallery_thumbnail_selected)

        # 3. Action Execution Card (Horizontally Centered, Fixed at Bottom of Column 2)
        act_card = SurfaceCard()
        act_main = QVBoxLayout()
        act_main.setContentsMargins(6, 4, 6, 4)
        act_main.setSpacing(6)

        # Extraction progress bar
        self.meta_left_progress_bar = QProgressBar()
        self.meta_left_progress_bar.setFixedHeight(4)
        self.meta_left_progress_bar.setTextVisible(False)
        self.meta_left_progress_bar.setRange(0, 100)
        self.meta_left_progress_bar.setValue(0)
        self.meta_left_progress_bar.setStyleSheet("""
            QProgressBar { background-color: #111422; border: none; border-radius: 2px; }
            QProgressBar::chunk { background-color: #3b82f6; border-radius: 2px; }
        """)
        act_main.addWidget(self.meta_left_progress_bar)

        act_layout = QHBoxLayout()
        act_layout.setContentsMargins(4, 2, 4, 2)
        act_layout.setSpacing(14)
        act_layout.addStretch()

        self.meta_btn_extract = QPushButton("Extract Metadata")
        self.meta_btn_extract.setIcon(LucideIcons.icon("file_text", color="#ffffff", size=18))
        self.meta_btn_extract.setIconSize(QSize(18, 18))
        self.meta_btn_extract.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_extract.setFixedWidth(180)
        self.meta_btn_extract.setFixedHeight(36)
        self.meta_btn_extract.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6;
                color: #ffffff;
                font-size: 12px;
                font-weight: 700;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #60a5fa; }
            QPushButton:disabled { background-color: #1e243b; color: #475569; }
        """)
        self.meta_btn_extract.clicked.connect(self._start_metadata_extraction)
        act_layout.addWidget(self.meta_btn_extract)

        self.meta_btn_cancel = QPushButton("Cancel")
        self.meta_btn_cancel.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_cancel.setFixedWidth(90)
        self.meta_btn_cancel.setFixedHeight(36)
        self.meta_btn_cancel.setEnabled(False)
        self.meta_btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #ef4444;
                color: #ffffff;
                font-size: 12px;
                font-weight: 700;
                padding: 10px 16px;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #f87171; }
            QPushButton:disabled { background-color: #1e243b; color: #475569; }
        """)
        self.meta_btn_cancel.clicked.connect(self._cancel_metadata_extraction)
        act_layout.addWidget(self.meta_btn_cancel)
        act_layout.addStretch()
        act_main.addLayout(act_layout)
        act_card.main_layout.addLayout(act_main)

        # Shims for removed Export Deliverables panel so existing worker callbacks remain robust
        self.meta_proc_status_lbl = QLabel("", self)
        self.meta_proc_status_lbl.hide()
        self.meta_quick_html_btn = QPushButton("", self)
        self.meta_quick_html_btn.hide()
        self.meta_quick_txt_btn = QPushButton("", self)
        self.meta_quick_txt_btn.hide()
        self.meta_quick_json_btn = QPushButton("", self)
        self.meta_quick_json_btn.hide()
        self.meta_btn_open_processed_folder = QPushButton("", self)
        self.meta_btn_open_processed_folder.hide()
        self.meta_quick_folder_btn = self.meta_btn_open_processed_folder
        self.meta_left_status_lbl = QLabel("", self)
        self.meta_left_status_lbl.hide()
        self.meta_left_count_lbl = QLabel("", self)
        self.meta_left_count_lbl.hide()
        self.meta_left_integrity_lbl = QLabel("", self)
        self.meta_left_integrity_lbl.hide()

        # Assembly: 1. Select File or Folder, 2. Dynamic Preview/Gallery Viewer (Expanding), 3. Action Panel (Fixed at Bottom)
        left_col_layout.addWidget(ev_card, 0)
        left_col_layout.addWidget(self.meta_preview_gallery, 1)
        left_col_layout.addWidget(act_card, 0)

        # ──────────────────────────────────────────────────────────────────────
        # RIGHT PANEL: EXTRACTION SETTINGS (TOP) & METADATA VIEWER (EXPANDED TO BOTTOM)
        # ──────────────────────────────────────────────────────────────────────
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(6, 0, 0, 0)
        right_layout.setSpacing(10)

        # 1. TOP CONTAINER: EXTRACTION SETTINGS
        cfg_card = SurfaceCard(title="Extraction Settings", icon="settings")

        cfg_checks_layout = QVBoxLayout()
        cfg_checks_layout.setSpacing(6)

        self.meta_cb_calc_hashes = QCheckBox("Calculate enabled file hashes (see Settings)")
        self.meta_cb_calc_hashes.setChecked(True)
        self.meta_cb_calc_hashes.setToolTip("Computes all cryptographic hashes supported by FINDEX")
        cfg_checks_layout.addWidget(self.meta_cb_calc_hashes)

        self.meta_cb_verify_integrity = QCheckBox("Verify file integrity (Compare hashes before & after)")
        self.meta_cb_verify_integrity.setChecked(True)
        self.meta_cb_verify_integrity.setToolTip("Calculates hashes before and after ExifTool execution to guarantee non-alteration")
        cfg_checks_layout.addWidget(self.meta_cb_verify_integrity)
        self.meta_cb_calc_hashes.toggled.connect(self._sync_metadata_integrity_option)
        self._sync_metadata_integrity_option()


        self.meta_cb_preserve_json = QCheckBox("Save raw ExifTool JSON data")
        self.meta_cb_preserve_json.setChecked(False)
        self.meta_cb_preserve_json.setToolTip(
            "Optional extra raw JSON file when JSON File (.json) is selected. "
            "HTML and text reports do not require a separate JSON file.")
        cfg_checks_layout.addWidget(self.meta_cb_preserve_json)

        cfg_card.main_layout.addLayout(cfg_checks_layout)

        # Deliverables & Output Directory Row
        deliv_out_row = QHBoxLayout()
        deliv_out_row.setSpacing(16)

        # Report formats options
        deliv_box = QVBoxLayout()
        deliv_box.setSpacing(4)
        fmt_lbl = QLabel("Report Formats:")
        fmt_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #8e98b0;")
        deliv_box.addWidget(fmt_lbl)

        fmt_row = QHBoxLayout()
        fmt_row.setSpacing(12)
        self.meta_cb_report_html = QCheckBox("HTML Report")
        self.meta_cb_report_html.setChecked(True)
        self.meta_cb_report_txt = QCheckBox("Text File (.txt)")
        self.meta_cb_report_txt.setChecked(False)
        self.meta_cb_report_json = QCheckBox("JSON File (.json)")
        self.meta_cb_report_json.setChecked(False)
        self.meta_cb_report_json.toggled.connect(self._sync_metadata_json_option)
        self._sync_metadata_json_option()
        fmt_row.addWidget(self.meta_cb_report_html)
        fmt_row.addWidget(self.meta_cb_report_txt)
        fmt_row.addWidget(self.meta_cb_report_json)
        fmt_row.addStretch()
        deliv_box.addLayout(fmt_row)
        deliv_out_row.addLayout(deliv_box, 1)

        # Output Folder
        out_box = QVBoxLayout()
        out_box.setSpacing(4)
        out_lbl = QLabel("Output Folder:")
        out_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #8e98b0;")
        out_box.addWidget(out_lbl)

        out_row = QHBoxLayout()
        out_row.setSpacing(6)
        self.meta_out_dir_edit = QLineEdit()
        self.meta_out_dir_edit.setPlaceholderText("Optional export destination folder...")
        self.meta_btn_browse_out = QPushButton("Browse")
        self.meta_btn_browse_out.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_browse_out.setStyleSheet("""
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 8px;
                color: #f1f5f9;
                padding: 6px 12px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
        """)
        self.meta_btn_browse_out.clicked.connect(self._browse_meta_out_dir)
        out_row.addWidget(self.meta_out_dir_edit)
        out_row.addWidget(self.meta_btn_browse_out)
        out_box.addLayout(out_row)
        deliv_out_row.addLayout(out_box, 1)

        cfg_card.main_layout.addLayout(deliv_out_row)

        # ExifTool Engine Status Bar
        engine_row = QHBoxLayout()
        engine_row.setSpacing(8)
        self.meta_engine_status_lbl = QLabel()
        self.meta_engine_status_lbl.setStyleSheet("font-size: 11px; font-weight: 600; padding: 4px 8px;")
        self._update_engine_status_pill()

        self.meta_btn_config_engine = QPushButton("Configure ExifTool")
        self.meta_btn_config_engine.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_config_engine.setStyleSheet("""
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #93c5fd;
                padding: 5px 12px;
                font-size: 11px;
                font-weight: 600;
            }
            QPushButton:hover { background-color: #28304e; border-color: #3b82f6; }
        """)
        self.meta_btn_config_engine.clicked.connect(self._open_configure_exiftool_dialog)

        engine_row.addWidget(self.meta_engine_status_lbl, 1)
        engine_row.addWidget(self.meta_btn_config_engine)
        cfg_card.main_layout.addLayout(engine_row)

        right_layout.addWidget(cfg_card)

        # 2. MAIN CONTAINER: METADATA VIEWER (Expands all the way down for maximum visibility)
        meta_view_card = SurfaceCard(title="Metadata Viewer", icon="search")
        meta_view_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Top Strip: Status, Spinner, Count Pill, Integrity Pill, Tags Count
        status_strip = QHBoxLayout()
        status_strip.setSpacing(8)

        self.meta_spinner = ProcessingSpinner(size=24)
        status_strip.addWidget(self.meta_spinner)

        self.meta_status_badge = QLabel("● Ready")
        self.meta_status_badge.setStyleSheet("font-size: 11px; font-weight: 700; color: #93c5fd; background: rgba(59, 130, 246, 0.12); border: 1px solid rgba(59, 130, 246, 0.25); border-radius: 6px; padding: 4px 10px;")
        status_strip.addWidget(self.meta_status_badge)

        self.meta_count_pill = QLabel("0 Files")
        self.meta_count_pill.setObjectName("CountPill")
        self.meta_count_pill.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0; background: rgba(14, 18, 29, 0.8); border: 1px solid #232b43; border-radius: 6px; padding: 4px 8px;")
        status_strip.addWidget(self.meta_count_pill)

        self.meta_integrity_pill = QLabel("Integrity: —")
        self.meta_integrity_pill.setObjectName("IntegrityPill")
        self.meta_integrity_pill.setStyleSheet("font-size: 11px; font-weight: 600; color: #8e98b0; background: rgba(14, 18, 29, 0.8); border: 1px solid #232b43; border-radius: 6px; padding: 4px 8px;")
        status_strip.addWidget(self.meta_integrity_pill)

        self.meta_tags_count_lbl = QLabel("Metadata Tags")
        self.meta_tags_count_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #93c5fd; padding-left: 6px;")
        self.btn_submode_meta = self.meta_tags_count_lbl
        status_strip.addWidget(self.meta_tags_count_lbl)

        status_strip.addStretch()
        meta_view_card.main_layout.addLayout(status_strip)

        # Micro progress bar
        self.meta_progress_bar = QProgressBar()
        self.meta_progress_bar.setFixedHeight(4)
        self.meta_progress_bar.setTextVisible(False)
        self.meta_progress_bar.setRange(0, 100)
        self.meta_progress_bar.setValue(0)
        self.meta_progress_bar.setStyleSheet("""
            QProgressBar { background-color: #111422; border: none; border-radius: 2px; }
            QProgressBar::chunk { background-color: #3b82f6; border-radius: 2px; }
        """)
        meta_view_card.main_layout.addWidget(self.meta_progress_bar)

        # Filter bar: Active file dropdown, live search, and category selector
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(8)

        self.meta_file_select_combo = QComboBox()
        self.meta_file_select_combo.currentIndexChanged.connect(self._on_meta_file_selected_combo)
        self.meta_file_select_combo.hide()
        filter_bar.addWidget(self.meta_file_select_combo)

        filter_bar.addStretch()

        self.meta_search_edit = QLineEdit()
        self.meta_search_edit.setPlaceholderText("Search tag or value...")
        self.meta_search_edit.setFixedWidth(220)
        self.meta_search_edit.textChanged.connect(self._filter_metadata_tree)
        filter_bar.addWidget(self.meta_search_edit)

        self.meta_cat_combo = QComboBox()
        self.meta_cat_combo.addItem("All Categories")
        for cat_name, _ in METADATA_CATEGORIES:
            self.meta_cat_combo.addItem(cat_name)
        self.meta_cat_combo.currentIndexChanged.connect(self._filter_metadata_tree)
        filter_bar.addWidget(self.meta_cat_combo)

        meta_view_card.main_layout.addLayout(filter_bar)

        # Full-height Metadata Tree
        self.meta_tree = QTreeWidget()
        self.meta_tree.setHeaderLabels(["Tag Name", "Value", "ExifTool Key"])
        self.meta_tree.setColumnWidth(0, 240)
        self.meta_tree.setColumnWidth(1, 360)
        self.meta_tree.setColumnWidth(2, 180)
        self.meta_tree.setAlternatingRowColors(False)
        self.meta_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.meta_tree.customContextMenuRequested.connect(self._on_meta_tree_context_menu)
        self.meta_tree.itemDoubleClicked.connect(lambda item, col: self._copy_selected_meta_value())
        meta_view_card.main_layout.addWidget(self.meta_tree, 1)

        # Tree Actions Toolbar
        tree_act_row = QHBoxLayout()
        tree_act_row.setSpacing(8)
        self.meta_btn_expand = QPushButton("Expand All")
        self.meta_btn_collapse = QPushButton("Collapse All")
        self.meta_btn_copy_val = QPushButton("Copy Value")

        for ab in (self.meta_btn_expand, self.meta_btn_collapse, self.meta_btn_copy_val):
            ab.setCursor(QCursor(Qt.PointingHandCursor))
            ab.setStyleSheet("""
                QPushButton {
                    background-color: #1e243b;
                    border: 1px solid #2e3654;
                    border-radius: 6px;
                    color: #f1f5f9;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background-color: #28304e;
                    border-color: #3b82f6;
                    color: #93c5fd;
                }
            """)
            tree_act_row.addWidget(ab)

        self.meta_btn_copy = self.meta_btn_copy_val
        self.meta_btn_expand.clicked.connect(self._expand_all_meta_tree)
        self.meta_btn_collapse.clicked.connect(self._collapse_all_meta_tree)
        self.meta_btn_copy_val.clicked.connect(self._copy_selected_meta_value)

        tree_act_row.addStretch()

        self.meta_btn_export_viewer = QPushButton("Export")
        self.meta_btn_export_html_viewer = self.meta_btn_export_viewer
        self.meta_btn_export_viewer.setIcon(LucideIcons.icon("download", color="#2563eb", size=14))
        self.meta_btn_export_viewer.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_export_viewer.setToolTip("Export metadata according to selected format extension(s) (.html, .txt, .json)")
        self._style_button_secondary(self.meta_btn_export_viewer, self.current_theme)
        self.meta_btn_export_viewer.clicked.connect(self._export_metadata_by_selected_format)
        tree_act_row.addWidget(self.meta_btn_export_viewer)

        self.meta_btn_view_raw_json = QPushButton("View Raw JSON")
        self.meta_btn_view_raw_json.setCursor(QCursor(Qt.PointingHandCursor))
        self.meta_btn_view_raw_json.setStyleSheet("""
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #3b82f6;
                color: #93c5fd;
                padding: 5px 14px;
                font-size: 11px;
                font-weight: 700;
                border-radius: 6px;
            }
            QPushButton:hover { background-color: #28304e; }
        """)
        self.meta_btn_view_raw_json.clicked.connect(self._open_raw_json_dialog)
        self.meta_btn_view_json = self.meta_btn_view_raw_json
        tree_act_row.addWidget(self.meta_btn_view_raw_json)

        meta_view_card.main_layout.addLayout(tree_act_row)
        right_layout.addWidget(meta_view_card, 1)

        # Shims and backward-compatible references for worker callbacks
        self.meta_stat_total = StatCardShim(self.meta_count_pill)
        self.meta_stat_success = StatCardShim(self.meta_count_pill)
        self.meta_stat_failed = StatCardShim(self.meta_count_pill)
        self.meta_stat_integrity = StatCardShim(self.meta_integrity_pill)
        self.meta_status_lbl = self.meta_status_badge
        self.meta_ticker_lbl = QLabel()

        # Splitter combining left and right panels with balanced distribution
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_col_widget)
        splitter.addWidget(right_container)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([390, 840])

        page_layout.addWidget(splitter)
        return page

    # ==========================================================================
    # METADATA EXPLORER CONTROLLER SLOTS & LOGIC
    # ==========================================================================
    def _update_submode_buttons(self):
        """Maintained for backward compatibility."""
        pass

    def _show_metadata_tab(self):
        """Maintained for backward compatibility."""
        pass

    def _show_preview_tab(self):
        """Opens selected evidence in the native viewer."""
        if hasattr(self, "meta_gallery_view"):
            self.meta_gallery_view._open_current_external()

    def _on_view_evidence_clicked(self):
        """Opens selected evidence in the native viewer."""
        if hasattr(self, "meta_gallery_view"):
            self.meta_gallery_view._open_current_external()

    def _update_engine_status_pill(self, refresh=False):
        """Display the engine state; probe only when configuration changes."""
        if refresh:
            self.exiftool_ver, self.exiftool_ok, self.exiftool_path = ExifToolEngine.get_version(self.custom_exiftool_path)
        ver, ok = self.exiftool_ver, self.exiftool_ok

        if ok:
            self.meta_engine_status_lbl.setText(f"ExifTool v{ver} Ready")
            self.meta_engine_status_lbl.setStyleSheet(("""
                color: #80cc93;
                font-weight: 700;
                font-size: 11px;
                background-color: rgba(128, 204, 147, 0.12);
                border: 1px solid rgba(128, 204, 147, 0.3);
                border-radius: 6px;
                padding: 4px 10px;
            """).replace("#80cc93", status_green(self)))
        else:
            self.meta_engine_status_lbl.setText("✕ ExifTool Not Found")
            self.meta_engine_status_lbl.setStyleSheet("""
                color: #f87171;
                font-weight: 700;
                font-size: 11px;
                background-color: rgba(239, 68, 68, 0.12);
                border: 1px solid rgba(239, 68, 68, 0.3);
                border-radius: 6px;
                padding: 4px 10px;
            """)

    def _open_configure_exiftool_dialog(self):
        dlg = ConfigureExifToolDialog(current_path=self.custom_exiftool_path, parent=self)
        dlg.configured.connect(self._on_exiftool_configured)
        dlg.exec()

    def _on_exiftool_configured(self, custom_path):
        self.custom_exiftool_path = custom_path.strip() if custom_path else None
        self._update_engine_status_pill(refresh=True)

    def _on_meta_target_mode_changed(self):
        is_folder = self.meta_rb_folder.isChecked()
        self.meta_cb_recursive.setEnabled(is_folder)
        if not is_folder:
            self.meta_cb_recursive.setChecked(False)
        if hasattr(self, "meta_btn_browse"):
            self.meta_btn_browse.setText("Browse")
            self.meta_btn_browse.setIcon(LucideIcons.icon("folder" if is_folder else "file_text", color="#2563eb", size=14))
            self.meta_btn_browse.setToolTip("Select a folder to inspect" if is_folder else "Select a file to inspect")

    def _on_meta_browse_clicked(self):
        if self.meta_rb_folder.isChecked():
            self._browse_meta_folder()
        else:
            self._browse_meta_file()

    def _browse_meta_file(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Select Evidence File for Metadata Extraction", "",
            "All Supported Evidence (*.*);;Images (*.jpg *.jpeg *.png *.tiff *.dng *.cr2 *.nef *.arw);;Audio/Video (*.mp4 *.mov *.mkv *.avi *.mp3 *.wav);;Documents (*.pdf *.docx *.xlsx *.pptx)"
        )
        if f:
            self.meta_rb_file.setChecked(True)
            self.meta_src_edit.setText(f)
            self._update_meta_evidence_display(f)

    def _browse_meta_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Select Folder for Metadata Extraction")
        if d:
            self.meta_rb_folder.setChecked(True)
            self.meta_src_edit.setText(d)
            self._update_meta_evidence_display(d)

    def _browse_meta_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Report Output Directory")
        if d:
            self.meta_out_dir_edit.setText(d)

    def _on_meta_path_dropped(self, dropped_path):
        if not dropped_path:
            return
        if os.path.isdir(dropped_path):
            self.meta_rb_folder.setChecked(True)
        else:
            self.meta_rb_file.setChecked(True)
        self.meta_src_edit.setText(dropped_path)
        self._update_meta_evidence_display(dropped_path)

    def _update_meta_evidence_display(self, target_text=""):
        p_str = self.meta_src_edit.text().strip()
        if not p_str or not os.path.exists(p_str):
            if hasattr(self, "meta_preview_gallery"):
                self.meta_preview_gallery.load_file(None)
            return

        p = Path(p_str)
        if p.is_dir():
            if not self.meta_rb_folder.isChecked():
                self.meta_rb_folder.blockSignals(True)
                self.meta_rb_folder.setChecked(True)
                self.meta_rb_folder.blockSignals(False)
            self.meta_cb_recursive.setEnabled(True)
            if hasattr(self, "meta_preview_gallery"):
                self.meta_preview_gallery.load_folder(p_str, recursive=self.meta_cb_recursive.isChecked())
            if not self.meta_out_dir_edit.text().strip():
                self.meta_out_dir_edit.setText(self.settings_store.value("default_output_dir", str(p / "Metadata_Reports")))
        elif p.is_file():
            if not self.meta_rb_file.isChecked():
                self.meta_rb_file.blockSignals(True)
                self.meta_rb_file.setChecked(True)
                self.meta_rb_file.blockSignals(False)
            self.meta_cb_recursive.setEnabled(False)
            if hasattr(self, "meta_preview_gallery"):
                self.meta_preview_gallery.load_file(p_str, record=getattr(self, "active_meta_record", None))
            if not self.meta_out_dir_edit.text().strip():
                self.meta_out_dir_edit.setText(self.settings_store.value("default_output_dir", str(p.parent / "Metadata_Reports")))

    def _on_gallery_thumbnail_selected(self, file_path):
        """Handles selection of a thumbnail in the folder gallery."""
        if hasattr(self, "all_meta_records") and self.all_meta_records:
            for rec in self.all_meta_records:
                if str(Path(rec.get("file_path", "")).resolve()) == str(Path(file_path).resolve()):
                    self.active_meta_record = rec
                    self._populate_metadata_tree(rec)
                    self._filter_metadata_tree()
                    if hasattr(self, "meta_file_select_combo"):
                        for i in range(self.meta_file_select_combo.count()):
                            item_rec = self.meta_file_select_combo.itemData(i)
                            if item_rec and item_rec.get("file_path") == rec.get("file_path"):
                                self.meta_file_select_combo.blockSignals(True)
                                self.meta_file_select_combo.setCurrentIndex(i)
                                self.meta_file_select_combo.blockSignals(False)
                                break
                    break

    def _sync_metadata_integrity_option(self):
        enabled = self.meta_cb_calc_hashes.isChecked()
        self.meta_cb_verify_integrity.setEnabled(enabled)
        if not enabled:
            self.meta_cb_verify_integrity.setChecked(False)
        self.meta_cb_verify_integrity.setToolTip(
            "Calculates hashes before and after ExifTool execution to verify non-alteration"
            if enabled else "Enable Calculate file hashes to verify integrity.")

    def _start_metadata_extraction(self):
        """Initiate metadata extraction workflow via background QThread."""
        src_path = self.meta_src_edit.text().strip()
        if not src_path or not os.path.exists(src_path):
            QMessageBox.warning(self, "Missing Evidence", "Please select a valid evidence file or folder.")
            return

        # Ensure ExifTool is operational
        ver, ok, exe_path = ExifToolEngine.get_version(self.custom_exiftool_path)
        if not ok:
            QMessageBox.critical(
                self, "ExifTool Engine Error",
                f"ExifTool extraction engine is unavailable:\n{exe_path}\n\nPlease click 'Configure Engine...' to locate exiftool.exe."
            )
            return

        # Collect files to process
        is_folder = os.path.isdir(src_path)
        file_paths = []

        if is_folder:
            recursive = self.meta_cb_recursive.isChecked()
            p = Path(src_path)
            if recursive:
                scan_errors = []
                for root, _, files in evidence_walk(p, onerror=lambda error: scan_errors.append(str(error))):
                    for fn in files:
                        file_paths.append(str(Path(root) / fn))
                if scan_errors:
                    QMessageBox.warning(self, "Folder Scan", "Some paths were skipped:\n" + "\n".join(scan_errors[:20]))
            else:
                for item in p.iterdir():
                    if item.is_file() and not linked(item):
                        file_paths.append(str(item))

            if not file_paths:
                QMessageBox.warning(self, "No Files Found", f"No files were found in the selected folder:\n{src_path}")
                return
        else:
            file_paths = [src_path]

        # Extraction runs in memory — files are saved strictly on-demand when the user chooses to export
        options = {
            "include_hashes": self.meta_cb_calc_hashes.isChecked(),
            "verify_integrity": self.meta_cb_verify_integrity.isChecked(),
            "preserve_raw_json": self._metadata_raw_export_requested(),
            "gen_html": False,
            "gen_txt": False,
            "gen_json": False,
            "output_dir": None,
            "auto_export": False,
            "custom_exiftool_path": self.custom_exiftool_path
        }

        # Reset UI controls
        self.meta_status_badge.setText("Extracting...")
        self.meta_status_badge.setStyleSheet("font-size: 11px; font-weight: 700; color: #fbbf24; background: rgba(245, 158, 11, 0.15); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 6px; padding: 4px 10px;")
        self.meta_count_pill.setText(f"0 / {len(file_paths)} Files")
        self.meta_integrity_pill.setText("Integrity: Evaluating...")
        self.meta_progress_bar.setRange(0, len(file_paths))
        self.meta_progress_bar.setValue(0)
        self.meta_progress_bar.show()

        if hasattr(self, "meta_left_status_lbl"):
            self.meta_left_status_lbl.setText("⏳ Extracting...")
            self.meta_left_status_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #60a5fa;")
            self.meta_left_count_lbl.setText(f"0 / {len(file_paths)} Files")
            self.meta_left_integrity_lbl.setText("Integrity: ...")
            self.meta_left_progress_bar.setRange(0, len(file_paths))
            self.meta_left_progress_bar.setValue(0)
            self.meta_proc_status_lbl.setText("Extracting metadata...")

        self.meta_btn_extract.setEnabled(False)
        self.meta_btn_cancel.setEnabled(True)
        self.meta_spinner.start()

        # Hide old report pills (folder button stays visible)
        for qb in (self.meta_quick_html_btn, self.meta_quick_txt_btn, self.meta_quick_json_btn):
            qb.hide()

        # Clear previous session items
        self.all_meta_records = []
        self.active_meta_record = None
        self.meta_file_select_combo.blockSignals(True)
        self.meta_file_select_combo.clear()
        self.meta_file_select_combo.hide()
        self.meta_file_select_combo.blockSignals(False)
        self.meta_tree.clear()

        # Launch background worker
        self.metadata_worker = MetadataExtractionWorker(file_paths, options, parent=self)
        self.metadata_worker.started_extraction.connect(self._on_meta_started)
        self.metadata_worker.file_started.connect(self._on_meta_file_started)
        self.metadata_worker.file_completed.connect(self._on_meta_file_completed)
        self.metadata_worker.file_failed.connect(self._on_meta_file_failed)
        self.metadata_worker.extraction_completed.connect(self._on_meta_completed)
        self.metadata_worker.extraction_cancelled.connect(self._on_meta_cancelled)
        self.metadata_worker.extraction_error.connect(self._on_meta_error)
        self.metadata_worker.start()

    def _on_meta_error(self, message):
        self.meta_spinner.stop()
        self.meta_btn_extract.setEnabled(True)
        self.meta_btn_cancel.setEnabled(False)
        self.meta_status_badge.setText("Extraction failed")
        QMessageBox.critical(self, "Metadata Extraction Error", message)

    def _cancel_metadata_extraction(self):
        if self.metadata_worker and self.metadata_worker.isRunning():
            self.metadata_worker.cancel()
            self.meta_status_badge.setText("Cancelling...")
            if hasattr(self, "meta_left_status_lbl"):
                self.meta_left_status_lbl.setText("Cancelling...")
                self.meta_left_status_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #f87171;")
            self.meta_btn_cancel.setEnabled(False)

    def _on_meta_started(self, total):
        self.meta_count_pill.setText(f"0 / {total} Files")
        self.meta_progress_bar.setRange(0, total)
        self.meta_progress_bar.setValue(0)
        if hasattr(self, "meta_left_count_lbl"):
            self.meta_left_count_lbl.setText(f"0 / {total} Files")
            self.meta_left_progress_bar.setRange(0, total)
            self.meta_left_progress_bar.setValue(0)
            self.meta_left_status_lbl.setText(f"⏳ Extracting 0/{total}...")
            self.meta_left_status_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #60a5fa;")

    def _on_meta_file_started(self, fpath, idx, total):
        fn = Path(fpath).name
        self.meta_status_badge.setText(f"Extracting ({idx}/{total})...")
        self.meta_progress_bar.setValue(idx - 1)
        if hasattr(self, "meta_left_status_lbl"):
            self.meta_left_status_lbl.setText(f"Extracting ({idx}/{total})...")
            self.meta_left_progress_bar.setValue(idx - 1)

    def _on_meta_file_completed(self, fpath, record, idx, total):
        self.meta_progress_bar.setValue(idx)
        self.all_meta_records.append(record)

        self.meta_count_pill.setText(f"{idx} / {total} Files")
        if hasattr(self, "meta_left_progress_bar"):
            self.meta_left_progress_bar.setValue(idx)
            self.meta_left_count_lbl.setText(f"{idx} / {total} Files")

        # Show combo if more than 1 file
        if total > 1:
            self.meta_file_select_combo.blockSignals(True)
            self.meta_file_select_combo.addItem(f"{Path(fpath).name} ({format_size(record.get('file_size', 0))})", userData=record)
            self.meta_file_select_combo.show()
            self.meta_file_select_combo.blockSignals(False)

        # Automatically display the first completed file
        if len(self.all_meta_records) == 1:
            self._display_record_metadata(record)

    def _on_meta_file_failed(self, fpath, error_msg, idx, total):
        self.meta_progress_bar.setValue(idx)
        if hasattr(self, "meta_left_progress_bar"):
            self.meta_left_progress_bar.setValue(idx)

    def _on_meta_completed(self, summary):
        self.meta_spinner.stop()
        self.meta_btn_extract.setEnabled(True)
        self.meta_btn_cancel.setEnabled(False)

        total = summary.get("total", 0)
        successful = summary.get("successful", 0)
        failed = summary.get("failed", 0)
        verified = summary.get("verified", 0)

        self.meta_status_badge.setText("Completed")
        self.meta_status_badge.setStyleSheet(("font-size: 11px; font-weight: 700; color: #80cc93; background: rgba(128, 204, 147, 0.15); border: 1px solid rgba(128, 204, 147, 0.3); border-radius: 6px; padding: 4px 10px;").replace("#80cc93", status_green(self)))

        self.meta_count_pill.setText(f"{successful} / {total} Files" if failed == 0 else f"{successful} ok, {failed} failed")

        # Log session activity for Overview telemetry
        target_path = self.meta_src_edit.text()
        self._log_session_activity(
            "Metadata Explorer",
            target_path,
            "Completed" if failed == 0 else "Warnings",
            f"Extracted metadata for {successful}/{total} file(s) via ExifTool {self.exiftool_ver}"
        )

        if self.meta_cb_calc_hashes.isChecked() and self.meta_cb_verify_integrity.isChecked():
            if verified == successful and successful > 0:
                self.meta_integrity_pill.setText("Integrity Verified")
                self.meta_integrity_pill.setStyleSheet(("font-size: 11px; font-weight: 700; color: #80cc93; background: rgba(128, 204, 147, 0.15); border: 1px solid rgba(128, 204, 147, 0.3); border-radius: 6px; padding: 4px 8px;").replace("#80cc93", status_green(self)))
            else:
                self.meta_integrity_pill.setText(f"Integrity: {verified}/{successful}")
                self.meta_integrity_pill.setStyleSheet("font-size: 11px; font-weight: 700; color: #f87171; background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: 6px; padding: 4px 8px;")
        else:
            self.meta_integrity_pill.setText("Integrity: N/A")

        # Sync Left Panel Processed Data Variables
        if hasattr(self, "meta_left_status_lbl"):
            self.meta_left_status_lbl.setText("Completed")
            self.meta_left_status_lbl.setStyleSheet(("font-size: 11px; font-weight: 700; color: #80cc93;").replace("#80cc93", status_green(self)))
            self.meta_left_count_lbl.setText(f"{successful} / {total} Files" if failed == 0 else f"{successful} ok, {failed} err")
            if self.meta_cb_calc_hashes.isChecked() and self.meta_cb_verify_integrity.isChecked():
                self.meta_left_integrity_lbl.setText("Verified" if (verified == successful and successful > 0) else f"Integrity: {verified}/{successful}")
                self.meta_left_integrity_lbl.setStyleSheet(("font-size: 11px; font-weight: 700; color: #80cc93;").replace("#80cc93", status_green(self)) if (verified == successful and successful > 0) else "font-size: 11px; font-weight: 700; color: #f87171;")
            else:
                self.meta_left_integrity_lbl.setText("Integrity: N/A")
            self.meta_proc_status_lbl.setText("Metadata extracted and ready. Select format to export:")

        # Show Deliverable Buttons for on-demand export
        self._populate_meta_report_buttons()

    def _on_meta_cancelled(self, processed_count):
        self.meta_spinner.stop()
        self.meta_btn_extract.setEnabled(True)
        self.meta_btn_cancel.setEnabled(False)
        self.meta_status_badge.setText("● Cancelled")
        self.meta_status_badge.setStyleSheet("font-size: 11px; font-weight: 700; color: #f87171; background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: 6px; padding: 4px 10px;")
        if hasattr(self, "meta_left_status_lbl"):
            self.meta_left_status_lbl.setText("● Cancelled")
            self.meta_left_status_lbl.setStyleSheet("font-size: 11px; font-weight: 700; color: #f87171;")
            self.meta_proc_status_lbl.setText("Extraction was cancelled.")

    def _on_meta_file_selected_combo(self, index):
        if index >= 0 and index < len(self.all_meta_records):
            record = self.meta_file_select_combo.itemData(index)
            if record:
                self._display_record_metadata(record)

    def _display_record_metadata(self, record):
        """Populate the metadata tree widget and update the preview widget."""
        self.active_meta_record = record
        self._populate_metadata_tree(record)
        self._filter_metadata_tree()
        if os.path.isdir(self.meta_src_edit.text().strip()):
            self.meta_preview_gallery.current_record = record
            self.meta_preview_gallery.selected_item_path = record.get("file_path")
            if self.meta_preview_gallery.current_mode == "folder":
                self.meta_preview_gallery._render_gallery_grid()
        else:
            self.meta_preview_widget.load_file(record.get("file_path"), record=record)

    def _populate_metadata_tree(self, record):
        self.meta_tree.clear()
        if not record:
            self.btn_submode_meta.setText("Metadata Tags")
            return

        categories = record.get("categories", {})
        total_tags = 0
        for cat_name, tags_list in categories.items():
            if not tags_list:
                continue

            total_tags += len(tags_list)

            # Top-level category header item
            cat_item = QTreeWidgetItem([f"{cat_name} ({len(tags_list)})", "", ""])
            cat_font = cat_item.font(0)
            cat_font.setBold(True)
            cat_item.setFont(0, cat_font)

            # High-contrast category background tint
            theme = getattr(self, "current_theme", "dark")
            if theme == "light":
                cat_item.setBackground(0, QColor("#f1f5f9"))
                cat_item.setBackground(1, QColor("#f1f5f9"))
                cat_item.setBackground(2, QColor("#f1f5f9"))
                cat_item.setForeground(0, QColor("#000000"))
            else:
                cat_item.setBackground(0, QColor("#1c2236"))
                cat_item.setBackground(1, QColor("#1c2236"))
                cat_item.setBackground(2, QColor("#1c2236"))
                cat_item.setForeground(0, QColor("#93c5fd"))

            cat_item.setData(0, Qt.UserRole, cat_name)

            for tag_dict in tags_list:
                tag_name = tag_dict.get("tag", "")
                raw_key = tag_dict.get("raw_key", "")
                val = str(tag_dict.get("value", ""))
                child = QTreeWidgetItem([tag_name, val, raw_key])
                child.setData(0, Qt.UserRole, cat_name)
                if theme == "light":
                    for c_col in (0, 1, 2):
                        child.setBackground(c_col, QColor("#ffffff"))
                        child.setForeground(c_col, QColor("#000000"))
                else:
                    for c_col in (0, 1, 2):
                        child.setBackground(c_col, QColor("#0e121d"))
                        child.setForeground(c_col, QColor("#f1f5f9"))
                cat_item.addChild(child)

            self.meta_tree.addTopLevelItem(cat_item)

        self.btn_submode_meta.setText(f"Metadata Tags ({total_tags})")

        # Expand major categories by default
        for i in range(self.meta_tree.topLevelItemCount()):
            top = self.meta_tree.topLevelItem(i)
            c_name = top.data(0, Qt.UserRole)
            if c_name in ("File Information", "EXIF", "GPS", "Image Information", "Document Metadata"):
                top.setExpanded(True)

    def _refresh_meta_tree_theme(self, theme):
        """Ensure all tree items strictly follow plain white/black scheme without colored rows."""
        if not hasattr(self, "meta_tree") or not self.meta_tree:
            return
        self.meta_tree.setAlternatingRowColors(False)
        for i in range(self.meta_tree.topLevelItemCount()):
            top = self.meta_tree.topLevelItem(i)
            if theme == "light":
                for col in (0, 1, 2):
                    top.setBackground(col, QColor("#f1f5f9"))
                    top.setForeground(col, QColor("#000000"))
            else:
                for col in (0, 1, 2):
                    top.setBackground(col, QColor("#1c2236"))
                    top.setForeground(col, QColor("#93c5fd"))

            for j in range(top.childCount()):
                child = top.child(j)
                if theme == "light":
                    for col in (0, 1, 2):
                        child.setBackground(col, QColor("#ffffff"))
                        child.setForeground(col, QColor("#000000"))
                else:
                    for col in (0, 1, 2):
                        child.setBackground(col, QColor("#0e121d"))
                        child.setForeground(col, QColor("#f1f5f9"))

    def _filter_metadata_tree(self):
        """Filter metadata tree based on search text and selected category."""
        search_txt = self.meta_search_edit.text().strip().lower()
        selected_cat = self.meta_cat_combo.currentText()

        for i in range(self.meta_tree.topLevelItemCount()):
            cat_item = self.meta_tree.topLevelItem(i)
            cat_name = cat_item.data(0, Qt.UserRole)

            # Category filter check
            if selected_cat != "All Categories" and cat_name != selected_cat:
                cat_item.setHidden(True)
                continue

            child_matches = 0
            for c_idx in range(cat_item.childCount()):
                child = cat_item.child(c_idx)
                tag_t = child.text(0).lower()
                val_t = child.text(1).lower()
                raw_t = child.text(2).lower()

                matches = (not search_txt) or (search_txt in tag_t) or (search_txt in val_t) or (search_txt in raw_t)
                child.setHidden(not matches)
                if matches:
                    child_matches += 1

            if search_txt:
                cat_item.setHidden(child_matches == 0)
                if child_matches > 0:
                    cat_item.setExpanded(True)
            else:
                cat_item.setHidden(False)

    def _expand_all_meta_tree(self):
        self.meta_tree.expandAll()

    def _collapse_all_meta_tree(self):
        self.meta_tree.collapseAll()

    def _copy_selected_meta_value(self):
        item = self.meta_tree.currentItem()
        if not item:
            return
        if item.childCount() > 0:
            val = item.text(0)
        else:
            val = item.text(1)
        QApplication.clipboard().setText(val)
        QMessageBox.information(self, "Copied", f"Copied to clipboard:\n{val[:150]}")

    def _on_meta_tree_context_menu(self, pos):
        item = self.meta_tree.itemAt(pos)
        if not item:
            return

        menu = QMenu(self)
        if item.childCount() == 0:
            act_copy_val = menu.addAction("Copy Value")
            act_copy_tag_val = menu.addAction("Copy Tag & Value")
            act_copy_key = menu.addAction("Copy ExifTool Raw Key")
            menu.addSeparator()
            act_view = menu.addAction("👁️ View File / Preview")

            action = menu.exec(self.meta_tree.viewport().mapToGlobal(pos))
            if action == act_copy_val:
                QApplication.clipboard().setText(item.text(1))
            elif action == act_copy_tag_val:
                QApplication.clipboard().setText(f"{item.text(0)}: {item.text(1)}")
            elif action == act_copy_key:
                QApplication.clipboard().setText(item.text(2))
            elif action == act_view:
                self._show_preview_tab()
        else:
            act_exp = menu.addAction("Expand Group")
            act_col = menu.addAction("Collapse Group")
            action = menu.exec(self.meta_tree.viewport().mapToGlobal(pos))
            if action == act_exp:
                item.setExpanded(True)
            elif action == act_col:
                item.setExpanded(False)

    def _open_raw_json_dialog(self):
        if not self.active_meta_record or not self.active_meta_record.get("raw_json"):
            QMessageBox.information(self, "No Raw JSON", "No raw ExifTool JSON is available for this file.")
            return

        dlg = RawJsonDialog(
            raw_json_text=self.active_meta_record.get("raw_json", "{}"),
            file_name=self.active_meta_record.get("file_name", "Evidence"),
            parent=self
        )
        dlg.exec()

    def _clear_meta_reports_layout(self):
        self._last_meta_html = None
        self._last_meta_txt = None
        self._last_meta_json = None
        self._last_meta_folder = None
        for qb in (self.meta_quick_html_btn, self.meta_quick_txt_btn, self.meta_quick_json_btn):
            qb.hide()
        if hasattr(self, "meta_proc_status_lbl"):
            self.meta_proc_status_lbl.setText("Reports will appear here after extraction.")

    def _get_active_meta_record(self):
        """Returns the currently viewed or first available metadata record."""
        if hasattr(self, "active_meta_record") and self.active_meta_record:
            return self.active_meta_record
        if hasattr(self, "current_meta_record") and self.current_meta_record:
            return self.current_meta_record
        if hasattr(self, "all_meta_records") and self.all_meta_records:
            return self.all_meta_records[0]
        return None

    def _write_meta_json_record(self, file_path, record):
        """Helper to write a forensic metadata record to JSON file."""
        fname = record.get("file_name", Path(record.get("file_path", "evidence")).name)
        integrity_info = record.get("integrity_info")
        include_hashes = bool(integrity_info and integrity_info.get("pre"))
        app_record = {
            "application": f"{APP_NAME} {APP_VERSION}",
            "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "extraction_engine": "ExifTool",
            "exiftool_version": record.get("exiftool_version", "Unknown"),
            "source_file": {
                "name": fname,
                "path": record.get("file_path", ""),
                "size_bytes": record.get("file_size", 0),
                "size_formatted": format_size(record.get("file_size", 0)),
            },
            "integrity_verification": {
                "hashes_enabled": include_hashes,
                "verified": integrity_info.get("verified", False) if integrity_info else None,
                "status": integrity_info.get("status", "NOT_REQUESTED") if integrity_info else "NOT_REQUESTED",
                "hashes_pre": integrity_info.get("pre", {}) if integrity_info else {},
                "hashes_post": integrity_info.get("post", {}) if integrity_info else {},
            },
            "warnings": record.get("warnings", []),
            "metadata_categories": record.get("categories", {}),
            "parsed_raw_metadata": record.get("parsed", {}),
            "raw_json": record.get("raw_json", "") if self._metadata_raw_export_requested() else None
        }
        Path(file_path).parent.mkdir(parents=True, exist_ok=True)
        raw_path = None
        if self._metadata_raw_export_requested():
            raw_stem = Path(file_path).stem.removesuffix("_metadata")
            raw_path = Path(file_path).with_name(raw_stem + "_exiftool_raw.json")
            write_report(raw_path, record.get("raw_json", "[]"))
        with new_report(file_path) as jf:
            json.dump(app_record, jf, ensure_ascii=False, indent=2)
        return str(raw_path) if raw_path else None

    @report_export_errors
    def _export_metadata_by_selected_format(self):
        """Exports metadata according to the format extensions selected at the top:
        HTML Report, Text File (.txt), and/or JSON File (.json)."""
        record = self._get_active_meta_record()
        all_records = getattr(self, "all_meta_records", [])
        if not record and not all_records:
            QMessageBox.information(self, "No Metadata", "Please extract metadata first before exporting.")
            return

        want_html = getattr(self, "meta_cb_report_html", None) and self.meta_cb_report_html.isChecked()
        want_txt = getattr(self, "meta_cb_report_txt", None) and self.meta_cb_report_txt.isChecked()
        want_json = getattr(self, "meta_cb_report_json", None) and self.meta_cb_report_json.isChecked()

        selected_formats = []
        if want_html:
            selected_formats.append("html")
        if want_txt:
            selected_formats.append("txt")
        if want_json:
            selected_formats.append("json")

        if not selected_formats:
            QMessageBox.warning(
                self,
                "No Format Selected",
                "Please select at least one format (HTML, Text, or JSON) under Report Formats at the top."
            )
            return

        is_batch = len(all_records) > 1
        records_to_export = all_records if is_batch else ([record] if record else [all_records[0]])
        records_to_export = [rec for rec in records_to_export if rec.get("success", True)]
        if not records_to_export:
            QMessageBox.warning(self, "Extraction Failed", "There are no successfully extracted records to export.")
            return

        # If single format and single file, use standard Save As dialog
        if len(selected_formats) == 1 and not is_batch:
            fmt = selected_formats[0]
            if fmt == "html":
                self._export_current_meta_html()
            elif fmt == "txt":
                self._export_current_meta_txt()
            elif fmt == "json":
                self._export_current_meta_json()
            return

        # If multiple formats or batch, ask for destination directory
        initial_dir = self.meta_out_dir_edit.text().strip()
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = getattr(self, "_last_meta_folder", "") or str(Path.home())

        fmt_names = ", ".join(f.upper() for f in selected_formats)
        target_dir = QFileDialog.getExistingDirectory(
            self,
            f"Select Export Destination Folder ({fmt_names})",
            initial_dir
        )
        if not target_dir:
            return

        saved_files = []
        for rec in records_to_export:
            fname = rec.get("file_name", Path(rec.get("file_path", "evidence")).name)
            sanitized_name = metadata_report_basename(rec)
            integrity_info = rec.get("integrity_info")
            include_hashes = bool(integrity_info and integrity_info.get("pre"))

            if want_html:
                out_html = os.path.join(target_dir, f"{sanitized_name}_Metadata.html")
                MetadataReportExporter.export_html(out_html, rec, integrity_info=integrity_info, include_hashes=include_hashes)
                saved_files.append(out_html)
                self._last_meta_html = out_html

            if want_txt:
                out_txt = os.path.join(target_dir, f"{sanitized_name}_Metadata.txt")
                MetadataReportExporter.export_txt(out_txt, rec, integrity_info=integrity_info, include_hashes=include_hashes)
                saved_files.append(out_txt)
                self._last_meta_txt = out_txt

            if want_json:
                out_json = os.path.join(target_dir, f"{sanitized_name}_metadata.json")
                raw_path = self._write_meta_json_record(out_json, rec)
                saved_files.append(out_json)
                if raw_path:
                    saved_files.append(raw_path)
                self._last_meta_json = out_json

        self._last_meta_folder = target_dir
        self.meta_out_dir_edit.setText(target_dir)
        # self.meta_btn_open_processed_folder.show()

        summary_msg = f"Successfully exported {len(saved_files)} report(s) in selected format(s):\n\n"
        summary_msg += f"Formats: {fmt_names}\n"
        summary_msg += f"Files Audited: {len(records_to_export)}\n"
        summary_msg += f"Saved to:\n{target_dir}"

        self._log_session_activity(
            "Metadata Explorer",
            f"{len(records_to_export)} file(s)",
            "Exported",
            f"Saved {len(saved_files)} report(s) ({', '.join(selected_formats)}) to {Path(target_dir).name}"
        )
        QMessageBox.information(self, "Export Complete", summary_msg)

    @report_export_errors
    def _export_current_meta_html(self):
        """Exports currently inspected metadata to a clean minimal white-theme HTML report."""
        record = self._get_active_meta_record()
        if not record or not record.get("success", True):
            QMessageBox.information(self, "No Metadata", "Please extract metadata first before exporting.")
            return
        fname = record.get("file_name", Path(record.get("file_path", "evidence")).name)
        sanitized_name = metadata_report_basename(record)
        initial_dir = self.meta_out_dir_edit.text().strip()
        default_file = f"{sanitized_name}_Metadata.html"
        if initial_dir and os.path.isdir(initial_dir):
            default_file = os.path.join(initial_dir, default_file)
        f, _ = QFileDialog.getSaveFileName(self, "Export Metadata HTML Report", default_file, "HTML Files (*.html)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            integrity_info = record.get("integrity_info")
            include_hashes = bool(integrity_info and integrity_info.get("pre"))
            MetadataReportExporter.export_html(f, record, integrity_info=integrity_info, include_hashes=include_hashes)
            self._last_meta_html = f
            self._last_meta_folder = str(Path(f).parent)
            # self.meta_btn_open_processed_folder.show()
            self._log_session_activity("Metadata Explorer", fname, "Exported", f"Saved HTML report to {Path(f).name}")
            QMessageBox.information(self, "Report Exported", f"Saved Metadata HTML Report:\n{f}")

    @report_export_errors
    def _export_current_meta_txt(self):
        """Exports currently inspected metadata to a clean structured plain text report."""
        record = self._get_active_meta_record()
        if not record or not record.get("success", True):
            QMessageBox.information(self, "No Metadata", "Please extract metadata first before exporting.")
            return
        fname = record.get("file_name", Path(record.get("file_path", "evidence")).name)
        sanitized_name = metadata_report_basename(record)
        initial_dir = self.meta_out_dir_edit.text().strip()
        default_file = f"{sanitized_name}_Metadata.txt"
        if initial_dir and os.path.isdir(initial_dir):
            default_file = os.path.join(initial_dir, default_file)
        f, _ = QFileDialog.getSaveFileName(self, "Export Metadata Text Report", default_file, "Text Files (*.txt)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            integrity_info = record.get("integrity_info")
            include_hashes = bool(integrity_info and integrity_info.get("pre"))
            MetadataReportExporter.export_txt(f, record, integrity_info=integrity_info, include_hashes=include_hashes)
            self._last_meta_txt = f
            self._last_meta_folder = str(Path(f).parent)
            # self.meta_btn_open_processed_folder.show()
            self._log_session_activity("Metadata Explorer", fname, "Exported", f"Saved Text report to {Path(f).name}")
            QMessageBox.information(self, "Report Exported", f"Saved Metadata Text Report:\n{f}")

    @report_export_errors
    def _export_current_meta_json(self):
        """Exports currently inspected metadata and raw JSON structures on-demand."""
        record = self._get_active_meta_record()
        if not record or not record.get("success", True):
            QMessageBox.information(self, "No Metadata", "Please extract metadata first before exporting.")
            return
        fname = record.get("file_name", Path(record.get("file_path", "evidence")).name)
        sanitized_name = metadata_report_basename(record)
        initial_dir = self.meta_out_dir_edit.text().strip()
        default_file = f"{sanitized_name}_metadata.json"
        if initial_dir and os.path.isdir(initial_dir):
            default_file = os.path.join(initial_dir, default_file)
        f, _ = QFileDialog.getSaveFileName(self, "Export Metadata JSON Record", default_file, "JSON Files (*.json)", options=QFileDialog.Option.DontConfirmOverwrite)
        if f:
            raw_path = self._write_meta_json_record(f, record)
            self._last_meta_json = f
            self._last_meta_folder = str(Path(f).parent)
            # self.meta_btn_open_processed_folder.show()
            self._log_session_activity("Metadata Explorer", fname, "Exported", f"Saved JSON record to {Path(f).name}")
            details = f"Saved Metadata JSON Record:\n{f}"
            if raw_path:
                details += f"\n\nRaw ExifTool JSON:\n{raw_path}"
            QMessageBox.information(self, "Record Exported", details)

    def _open_quick_html(self):
        self._export_current_meta_html()

    def _open_quick_txt(self):
        self._export_current_meta_txt()

    def _open_quick_json(self):
        self._export_current_meta_json()

    def _open_quick_folder(self):
        folder = getattr(self, "_last_meta_folder", None)
        if folder and os.path.exists(folder):
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
        else:
            QMessageBox.information(
                self, "Export Folder",
                "No export folder has been selected yet.\nReports are saved only when you click Export."
            )

    def _populate_meta_report_buttons(self, report_paths=None, output_dir=None):
        """No-op: Exporting is initiated on-demand via the dedicated Column 3 Export button."""
        pass


# ==============================================================================
# SECTION 8: APPLICATION ENTRY POINT
# ==============================================================================

    def _set_hash_controls_enabled(self, enabled):
        for widget in (self.hash_btn_add_files, self.hash_btn_add_folder, self.hash_btn_load_checksum,
                       self.hash_btn_clear, self.hash_btn_set_target, self.hash_btn_paste,
                       self.hash_btn_clear_tgt, self.hash_algo_combo, self.target_hash_edit):
            widget.setEnabled(enabled)
        self.hash_table.setEnabled(enabled)

    def _save_raw_meta_payload(self, directory, record):
        if not self._metadata_raw_export_requested():
            return None
        target = Path(directory) / (metadata_report_basename(record) + "_exiftool_raw.json")
        write_report(target, record.get("raw_json", "[]"))
        return str(target)

    def _metadata_raw_export_requested(self):
        return self.meta_cb_report_json.isChecked() and self.meta_cb_preserve_json.isChecked()

    def _sync_metadata_json_option(self):
        self.meta_cb_preserve_json.setEnabled(self.meta_cb_report_json.isChecked())

    def _default_report_file(self, filename):
        directory = self.settings_store.value("default_output_dir", "")
        return str(Path(directory) / filename) if directory else filename

    def _report_evidence_targets(self):
        paths = [item.get('path') for item in getattr(self, 'hash_items', [])]
        records = list(getattr(self, 'all_meta_records', []))
        active = self._get_active_meta_record()
        if active:
            records.append(active)
        paths.extend(record.get('file_path') for record in records)
        roots = []
        for name in ('inv_src_edit', 'meta_src_edit'):
            field = getattr(self, name, None)
            value = field.text().strip() if field else ''
            if value:
                (roots if Path(value).is_dir() else paths).append(value)
        return [path for path in paths if path], roots

    def _confirm_report_replace(self, path):
        return QMessageBox.question(self, 'Replace Existing File?',
            f'A file already exists:\n{path}\n\nReplace it with the new report?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes

    def _approve_report_replacements(self, paths):
        protected, roots = self._report_evidence_targets()
        approved = {}
        try:
            with report_write_scope(protected_paths=protected, protected_roots=roots) as policy:
                for path in paths:
                    target = Path(path).absolute()
                    if target.exists() or target.is_symlink():
                        identity = check_report_target(target, policy)
                        if not self._confirm_report_replace(target):
                            return None
                        approved[report_key(target)] = identity
        except (OSError, ValueError) as error:
            QMessageBox.warning(self, 'Cannot Save Report', str(error))
            return None
        return approved

    def closeEvent(self, event):
        workers = [worker for worker in (self.inventory_worker, self.hash_worker, self.metadata_worker)
                   if worker is not None and worker.isRunning()]
        if workers:
            for worker in workers:
                worker.cancel()
            event.ignore()
            self.statusBar().showMessage("Stopping active work before closing...")
            QTimer.singleShot(100, self.close)
            return
        event.accept()


def main():
    # Use FINDEX identity for Windows taskbar grouping and Alt+Tab.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("FINDEX.FileIntelligenceExamination")
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyle("Fusion")

    # Use the primary FINDEX favicon for desktop application icons.
    assets = ensure_ui_assets()
    logo_file = assets.get("logo")
    if logo_file and os.path.exists(logo_file):
        app.setWindowIcon(application_icon())

    # High-quality clean Segoe UI font
    font = QFont("Segoe UI", 10)
    font.setStyleHint(QFont.SansSerif)
    app.setFont(font)

    window = FileInventoryProApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--check-excel":
        from excel_diagnostics import check_excel_exports
        sys.exit(check_excel_exports(sys.modules[__name__], sys.argv[2]))
    else:
        main()
