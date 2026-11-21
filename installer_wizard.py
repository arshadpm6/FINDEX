import sys
import os
import zipfile
import subprocess
import json
import threading
import tempfile
import hashlib
from pathlib import Path
from release_integrity import verify_payload

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QCheckBox,
    QProgressBar, QStackedWidget, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QIcon, QPixmap, QFont, QColor

def get_bundle_dir():
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            return Path(sys._MEIPASS).resolve()
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def find_payload_zip():
    b_dir = get_bundle_dir()
    candidates = [
        b_dir / "app_payload.zip",
        b_dir / "dist" / "app_payload.zip",
        Path(__file__).parent / "dist" / "app_payload.zip",
        Path(__file__).parent / "app_payload.zip",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def validate_destination(dest_dir):
    requested = Path(dest_dir).expanduser().absolute()
    if requested.is_symlink() or requested.is_junction():
        raise ValueError("The installation folder must not be a linked directory.")
    destination = requested.resolve()
    if destination.name.casefold() not in ("findex", "file inventory pro"):
        raise ValueError("Choose a dedicated folder named 'FINDEX'. Existing File Inventory Pro folders are accepted for upgrades.")
    if destination.exists():
        if not destination.is_dir() or destination.is_symlink() or destination.is_junction():
            raise ValueError("The installation folder must be a regular directory.")
        if any(destination.iterdir()) and not (destination / "install_manifest.json").is_file() and not (destination / "File Inventory Pro.exe").is_file() and not (destination / "FINDEX.exe").is_file():
            raise ValueError("This folder contains unrelated files. Choose an empty application folder.")
    return destination


class ExtractWorker(QThread):
    progress = Signal(int, int, str)
    finished = Signal(bool, str)

    def __init__(self, zip_path, dest_dir, make_desktop, make_start, make_reg):
        super().__init__()
        self.zip_path = zip_path
        self.dest_dir = Path(dest_dir)
        self.make_desktop = make_desktop
        self.make_start = make_start
        self.make_reg = make_reg
        self._cancel_event = threading.Event()

    def run(self):
        extraction_started = False
        try:
            self.dest_dir = validate_destination(self.dest_dir)
            installed = set()
            manifest_path = self.dest_dir / "install_manifest.json"
            for reserved in (manifest_path, self.dest_dir / 'uninstall.ps1', self.dest_dir / 'uninstall.bat'):
                if reserved.is_symlink() or reserved.is_junction():
                    raise ValueError('Refusing linked installer control files')
            if manifest_path.exists():
                installed.update(json.loads(manifest_path.read_text(encoding="utf-8")).get("files", []))
            owned = {str(Path(name)).casefold() for name in installed}
            # Legacy installs without a manifest may replace only the known executable.
            if not manifest_path.exists():
                owned.update(('findex.exe', 'file inventory pro.exe'))
            # Check the entire archive before creating or replacing any application file.
            with zipfile.ZipFile(self.zip_path, "r") as archive:
                payload = verify_payload(archive, Path(self.zip_path).with_suffix('.manifest.json'), self._cancel_event)
                for info in archive.infolist():
                    relative = Path(info.filename)
                    requested = self.dest_dir / relative
                    target = requested.resolve()
                    if not target.is_relative_to(self.dest_dir) or target == self.dest_dir:
                        raise ValueError('Payload path leaves the installation directory')
                    probe = requested
                    while probe != self.dest_dir:
                        if probe.is_symlink() or probe.is_junction():
                            raise ValueError('Refusing a linked installation path: ' + info.filename)
                        probe = probe.parent
                    if target.exists() and not info.is_dir() and str(relative).casefold() not in owned:
                        raise ValueError('An unrelated file would be overwritten: ' + info.filename)
            self.dest_dir.mkdir(parents=True, exist_ok=True)
            extraction_started = True
            # Mark partial installations so an interrupted extraction can be retried.
            manifest_path.write_text(json.dumps({"files": sorted(installed)}), encoding="utf-8")
            with zipfile.ZipFile(self.zip_path, "r") as archive:
                entries = archive.infolist()
                for i, info in enumerate(entries, 1):
                    if self._cancel_event.is_set():
                        raise InterruptedError("Installation cancelled. You can run Setup again to finish.")
                    relative = Path(info.filename)
                    target = (self.dest_dir / relative).resolve()
                    if relative.is_absolute() or ".." in relative.parts or not target.is_relative_to(self.dest_dir) or target == self.dest_dir:
                        raise ValueError("Unsafe path in application payload: " + info.filename)
                    if info.is_dir():
                        target.mkdir(parents=True, exist_ok=True)
                    else:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        temporary = None
                        try:
                            digest = hashlib.sha256()
                            with archive.open(info) as source, tempfile.NamedTemporaryFile(
                                    dir=target.parent, prefix='.findex-install-', delete=False) as output:
                                temporary = Path(output.name)
                                for chunk in iter(lambda: source.read(1024 * 1024), b''):
                                    if self._cancel_event.is_set():
                                        raise InterruptedError('Installation cancelled. Run Setup again to finish.')
                                    digest.update(chunk)
                                    output.write(chunk)
                            expected = payload['files'][info.filename]
                            if temporary.stat().st_size != expected['size'] or digest.hexdigest() != expected['sha256']:
                                raise ValueError('Extracted file failed integrity verification')
                            os.replace(temporary, target)
                        finally:
                            if temporary is not None:
                                temporary.unlink(missing_ok=True)
                        installed.add(str(relative))
                    self.progress.emit(i, len(entries), info.filename)
            main_exe = self.dest_dir / "FINDEX.exe"
            if not main_exe.is_file():
                raise ValueError("Application executable is missing from the payload")
            ico_file = self.dest_dir / "_internal" / "resources" / "branding" / "icon.ico"
            if not ico_file.exists():
                ico_file = self.dest_dir / "resources" / "branding" / "icon.ico"
            if self.make_desktop or self.make_start:
                self._remove_legacy_shortcuts()
            if self.make_desktop:
                self._create_shortcut(self._desktop_directory() / "FINDEX.lnk", main_exe, self.dest_dir, ico_file)
            if self.make_start:
                start_dir = Path(os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs\FINDEX"))
                start_dir.mkdir(parents=True, exist_ok=True)
                self._create_shortcut(start_dir / "FINDEX.lnk", main_exe, self.dest_dir, ico_file)
            installed.update(("uninstall.bat", "uninstall.ps1", "install_manifest.json"))
            manifest_path.write_text(json.dumps({"files": sorted(installed), "desktop": self.make_desktop,
                "start_menu": self.make_start, "registered": self.make_reg}, indent=2), encoding="utf-8")
            uninst_bat = self.dest_dir / "uninstall.bat"
            self._write_uninstaller(uninst_bat, self.dest_dir)
            if self.make_reg:
                self._write_registry(self.dest_dir, main_exe, ico_file, uninst_bat)
            self.finished.emit(True, "Installation completed successfully.")
        except Exception as exc:
            if extraction_started and "installed" in locals() and self.dest_dir.is_dir():
                try:
                    (self.dest_dir / "install_manifest.json").write_text(json.dumps({"files": sorted(installed),
                        "desktop": self.make_desktop, "start_menu": self.make_start, "registered": self.make_reg}), encoding="utf-8")
                except OSError:
                    pass
            self.finished.emit(False, str(exc))

    def cancel(self):
        self._cancel_event.set()

    @staticmethod
    def _desktop_directory():
        import ctypes
        buffer = ctypes.create_unicode_buffer(32768)
        result = ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer)
        if result != 0:
            raise OSError("Unable to locate the Desktop folder")
        return Path(buffer.value)

    def _remove_legacy_shortcuts(self):
        def quote(value):
            return "'" + str(value).replace("'", "''") + "'"
        paths = [self._desktop_directory() / "File Inventory Pro.lnk",
                 Path(os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs\File Inventory Pro\File Inventory Pro.lnk"))]
        targets = [self.dest_dir / "File Inventory Pro.exe", self.dest_dir / "FINDEX.exe"]
        command = "$ErrorActionPreference='Stop'; $shell=New-Object -ComObject WScript.Shell; "
        command += "$targets=@(" + ",".join(quote(path) for path in targets) + "); "
        command += "foreach($path in @(" + ",".join(quote(path) for path in paths) + ")) { "
        command += "if(Test-Path -LiteralPath $path) { $link=$shell.CreateShortcut($path); "
        command += "if($targets -contains $link.TargetPath) { Remove-Item -LiteralPath $path -Force } } }"
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                       creationflags=0x08000000, check=True, capture_output=True, text=True)

    def _create_shortcut(self, lnk_path, target_exe, work_dir, icon_path):
        def quote(value):
            return "'" + str(value).replace("'", "''") + "'"
        command = (
            "$ErrorActionPreference='Stop'; $ws=New-Object -ComObject WScript.Shell; "
            f"$s=$ws.CreateShortcut({quote(lnk_path)}); $s.TargetPath={quote(target_exe)}; "
            f"$s.WorkingDirectory={quote(work_dir)}; $s.IconLocation={quote(str(icon_path) + ',0')}; "
            "$s.Description='FINDEX'; $s.Save()"
        )
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                       creationflags=0x08000000, check=True, capture_output=True, text=True)

    def _write_uninstaller(self, uninst_bat, dest_dir):
        script = "$ErrorActionPreference = 'Stop'\n$root = [System.IO.Path]::GetFullPath($PSScriptRoot)\nif ([System.IO.Path]::GetFileName($root) -notin @('FINDEX', 'File Inventory Pro')) { throw 'Invalid application directory' }\n$manifestPath = Join-Path $root 'install_manifest.json'\n$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json\n$prefix = $root.TrimEnd('\\') + '\\'\n$targets = @()\nforeach ($relative in $manifest.files) {\n    if ([System.IO.Path]::IsPathRooted($relative) -or ($relative -split '[\\\\/]' -contains '..')) { throw 'Invalid installed path' }\n    $target = [System.IO.Path]::GetFullPath((Join-Path $root $relative))\n    if (-not $target.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'Path leaves application directory' }\n    $probe = $target\n    while ($probe -and $probe.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {\n        if (Test-Path -LiteralPath $probe) {\n            $entry = Get-Item -LiteralPath $probe -Force\n            if ($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { throw 'Refusing a linked installed path' }\n        }\n        $probe = [System.IO.Path]::GetDirectoryName($probe)\n    }\n    $targets += $target\n}\n# Delete only recorded files. Never recursively remove the installation folder.\n$directories = @()\nforeach ($target in $targets) {\n    if ([System.IO.File]::Exists($target)) { [System.IO.File]::Delete($target) }\n    $parent = [System.IO.Path]::GetDirectoryName($target)\n    while ($parent -and $parent.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {\n        $directories += $parent\n        $parent = [System.IO.Path]::GetDirectoryName($parent)\n    }\n}\nforeach ($directory in ($directories | Sort-Object -Unique | Sort-Object Length -Descending)) {\n    try { [System.IO.Directory]::Delete($directory, $false) } catch { }\n}\nif ($manifest.desktop) { [System.IO.File]::Delete((Join-Path ([Environment]::GetFolderPath('Desktop')) 'FINDEX.lnk')) }\nif ($manifest.start_menu) {\n    $start = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\\Windows\\Start Menu\\Programs\\FINDEX'\n    [System.IO.File]::Delete((Join-Path $start 'FINDEX.lnk'))\n    try { [System.IO.Directory]::Delete($start, $false) } catch { }\n}\nif ($manifest.registered) { Remove-Item -LiteralPath 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\FileInventoryPro' -ErrorAction SilentlyContinue }\ntry { [System.IO.Directory]::Delete($root, $false) } catch { Write-Host 'Uninstalled. Unrelated files were retained.' }\n"
        (Path(dest_dir) / "uninstall.ps1").write_text(script, encoding="utf-8-sig")
        Path(uninst_bat).write_text('@echo off\npowershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"\n', encoding="ascii")

    def _write_registry(self, dest_dir, main_exe, ico_file, uninst_bat):
        import winreg
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\FileInventoryPro"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                winreg.SetValueEx(k, "DisplayName", 0, winreg.REG_SZ, "FINDEX")
                winreg.SetValueEx(k, "DisplayVersion", 0, winreg.REG_SZ, "3.0.1")
                winreg.SetValueEx(k, "Publisher", 0, winreg.REG_SZ, "Antigravity Forensics")
                winreg.SetValueEx(k, "InstallLocation", 0, winreg.REG_SZ, str(dest_dir))
                winreg.SetValueEx(k, "DisplayIcon", 0, winreg.REG_SZ, str(ico_file if ico_file.exists() else main_exe))
                winreg.SetValueEx(k, "UninstallString", 0, winreg.REG_SZ, f'cmd.exe /c "{str(uninst_bat)}"')
                winreg.SetValueEx(k, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(k, "NoRepair", 0, winreg.REG_DWORD, 1)
        except OSError as exc:
            raise OSError('Application files installed, but Windows Installed Apps registration failed. '
                          'Retry Setup with registration unchecked.') from exc


class InstallerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FINDEX v3.0.1 - Setup Wizard")
        self.setFixedSize(540, 560)

        icon_path = get_bundle_dir() / "resources" / "branding" / "favicon.png"
        if not icon_path.exists():
            icon_path = Path(__file__).parent / "resources" / "branding" / "favicon.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.dest_path = os.path.expandvars(r"%LOCALAPPDATA%\Programs\FINDEX")
        legacy_dir = Path(os.path.expandvars(r"%LOCALAPPDATA%\Programs\File Inventory Pro"))
        if (legacy_dir / "install_manifest.json").is_file() or (legacy_dir / "File Inventory Pro.exe").is_file():
            self.dest_path = str(legacy_dir)

        self._build_ui()
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0b0e17;
                color: #f1f5f9;
            }
            QLabel {
                color: #e2e8f0;
                font-size: 12px;
            }
            QLineEdit {
                background-color: #161b2b;
                border: 1px solid #232b43;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 6px 10px;
                font-size: 11px;
            }
            QLineEdit:focus {
                border-color: #3b82f6;
            }
            QPushButton {
                background-color: #1e243b;
                border: 1px solid #2e3654;
                border-radius: 6px;
                color: #f1f5f9;
                padding: 6px 14px;
                font-weight: 600;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #28304e;
                border-color: #3b82f6;
            }
            QPushButton#btn_primary {
                background-color: #3b82f6;
                color: #ffffff;
                font-weight: 700;
                border: none;
                padding: 8px 18px;
            }
            QPushButton#btn_primary:hover {
                background-color: #60a5fa;
            }
            QCheckBox {
                color: #cbd5e1;
                font-size: 11px;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #334155;
                border-radius: 4px;
                background-color: #161b2b;
            }
            QCheckBox::indicator:checked {
                background-color: #3b82f6;
                border-color: #60a5fa;
            }
            QProgressBar {
                background-color: #161b2b;
                border: 1px solid #232b43;
                border-radius: 6px;
                text-align: center;
                color: #ffffff;
                font-size: 11px;
                font-weight: 600;
            }
            QProgressBar::chunk {
                background-color: #3b82f6;
                border-radius: 5px;
            }
        """)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(16)

        # Header Banner
        header = QFrame()
        header.setStyleSheet("background-color: #111422; border: 1px solid #1c2236; border-radius: 8px; padding: 12px;")
        h_layout = QHBoxLayout(header)
        h_layout.setContentsMargins(10, 8, 10, 8)
        h_layout.setSpacing(14)

        icon_lbl = QLabel()
        ico_file = get_bundle_dir() / "resources" / "branding" / "favicon.png"
        if not ico_file.exists():
            ico_file = Path(__file__).parent / "resources" / "branding" / "favicon.png"
        if ico_file.exists():
            icon_lbl.setPixmap(QPixmap(str(ico_file)).scaled(44, 44, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        h_layout.addWidget(icon_lbl)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title_lbl = QLabel("FINDEX")
        title_lbl.setStyleSheet("font-size: 16px; font-weight: 800; color: #f1f5f9;")
        sub_lbl = QLabel("Forensic File Details, Cryptographic Hashes & ExifTool Integration")
        sub_lbl.setStyleSheet("font-size: 11px; color: #94a3b8;")
        ver_lbl = QLabel("Version 3.0.1 • Windows x64 Standalone Deployment")
        ver_lbl.setStyleSheet("font-size: 10px; color: #60a5fa; font-weight: 700;")
        title_col.addWidget(title_lbl)
        title_col.addWidget(sub_lbl)
        title_col.addWidget(ver_lbl)
        h_layout.addLayout(title_col, 1)

        main_layout.addWidget(header)

        # Stacked Pages: 0: Options, 1: Installing, 2: Completed
        self.stack = QStackedWidget()

        # Page 0: Options
        page_opt = QWidget()
        p0_layout = QVBoxLayout(page_opt)
        p0_layout.setContentsMargins(0, 4, 0, 0)
        p0_layout.setSpacing(12)

        p0_lbl = QLabel("Select Destination Folder:")
        p0_lbl.setStyleSheet("font-weight: 700; color: #e2e8f0;")
        p0_layout.addWidget(p0_lbl)

        path_row = QHBoxLayout()
        self.edit_dest = QLineEdit(self.dest_path)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self._browse_dest)
        path_row.addWidget(self.edit_dest, 1)
        path_row.addWidget(btn_browse)
        p0_layout.addLayout(path_row)

        payload = find_payload_zip()
        required_mb = None
        if payload:
            try:
                with zipfile.ZipFile(payload) as archive:
                    required_mb = sum(info.file_size for info in archive.infolist()) / (1024 * 1024)
            except (OSError, zipfile.BadZipFile):
                pass
        space = f"~{required_mb:.0f} MB" if required_mb is not None else "varies by package"
        desc_lbl = QLabel(f"Space required: {space} | Per-user installation")
        desc_lbl.setStyleSheet("font-size: 10px; color: #64748b;")
        p0_layout.addWidget(desc_lbl)

        p0_layout.addSpacing(6)
        opt_hdr = QLabel("Installation Options:")
        opt_hdr.setStyleSheet("font-weight: 700; color: #e2e8f0;")
        p0_layout.addWidget(opt_hdr)

        self.cb_desktop = QCheckBox("Create Desktop Shortcut")
        self.cb_desktop.setChecked(True)
        self.cb_start = QCheckBox("Create Start Menu Shortcut")
        self.cb_start.setChecked(True)
        self.cb_reg = QCheckBox("Register in Windows Settings / Installed Apps (Easy Uninstall)")
        self.cb_reg.setChecked(True)

        p0_layout.addWidget(self.cb_desktop)
        p0_layout.addWidget(self.cb_start)
        p0_layout.addWidget(self.cb_reg)
        p0_layout.addStretch()

        self.stack.addWidget(page_opt)

        # Page 1: Progress
        page_prog = QWidget()
        p1_layout = QVBoxLayout(page_prog)
        p1_layout.setContentsMargins(0, 16, 0, 0)
        p1_layout.setSpacing(10)

        self.prog_status = QLabel("Extracting application files...")
        self.prog_status.setStyleSheet("font-weight: 700; color: #60a5fa;")
        p1_layout.addWidget(self.prog_status)

        self.prog_bar = QProgressBar()
        self.prog_bar.setFixedHeight(18)
        self.prog_bar.setRange(0, 100)
        self.prog_bar.setValue(0)
        p1_layout.addWidget(self.prog_bar)

        self.prog_file = QLabel("")
        self.prog_file.setStyleSheet("font-size: 10px; color: #64748b;")
        p1_layout.addWidget(self.prog_file)
        p1_layout.addStretch()

        self.stack.addWidget(page_prog)

        # Page 2: Finished
        page_fin = QWidget()
        p2_layout = QVBoxLayout(page_fin)
        p2_layout.setContentsMargins(0, 16, 0, 0)
        p2_layout.setSpacing(12)

        fin_title = QLabel("✓ Installation Complete!")
        fin_title.setStyleSheet("font-size: 16px; font-weight: 800; color: #34d399;")
        p2_layout.addWidget(fin_title)

        self.fin_desc = QLabel("FINDEX v3.0.1 has been successfully installed on your computer.")
        self.fin_desc.setStyleSheet("color: #cbd5e1;")
        p2_layout.addWidget(self.fin_desc)

        self.cb_launch = QCheckBox("Launch FINDEX now")
        self.cb_launch.setChecked(True)
        p2_layout.addWidget(self.cb_launch)
        p2_layout.addStretch()

        self.stack.addWidget(page_fin)

        main_layout.addWidget(self.stack, 1)

        # Bottom Action Bar
        bottom_row = QHBoxLayout()
        bottom_row.addStretch()

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.close)
        bottom_row.addWidget(self.btn_cancel)

        self.btn_action = QPushButton("Install")
        self.btn_action.setObjectName("btn_primary")
        self.btn_action.clicked.connect(self._on_action_clicked)
        bottom_row.addWidget(self.btn_action)

        main_layout.addLayout(bottom_row)

    def _browse_dest(self):
        chosen = QFileDialog.getExistingDirectory(self, "Select Install Destination", self.edit_dest.text())
        if chosen:
            self.edit_dest.setText(os.path.join(chosen, "FINDEX"))

    def _on_action_clicked(self):
        idx = self.stack.currentIndex()
        if idx == 0:
            dest = self.edit_dest.text().strip()
            if not dest:
                return
            try:
                self.dest_path = str(validate_destination(dest))
            except ValueError as exc:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "Invalid Destination", str(exc))
                return
            payload = find_payload_zip()
            if not payload or not payload.exists():
                self.prog_status.setText("Error: app_payload.zip not found in installer package.")
                self.stack.setCurrentIndex(1)
                self.btn_action.setEnabled(False)
                return

            self.stack.setCurrentIndex(1)
            self.btn_action.setEnabled(False)
            self.btn_cancel.setEnabled(False)

            self.worker = ExtractWorker(
                payload, self.dest_path,
                self.cb_desktop.isChecked(),
                self.cb_start.isChecked(),
                self.cb_reg.isChecked()
            )
            self.worker.progress.connect(self._on_progress)
            self.worker.finished.connect(self._on_finished)
            self.worker.start()

        elif idx == 2:
            if self.cb_launch.isChecked():
                exe_path = os.path.join(self.dest_path, "FINDEX.exe")
                if os.path.exists(exe_path):
                    subprocess.Popen([exe_path], cwd=self.dest_path)
            self.close()

    def _on_progress(self, current, total, filename):
        self.prog_bar.setRange(0, total)
        self.prog_bar.setValue(current)
        self.prog_status.setText(f"Installing files... ({current}/{total})")
        self.prog_file.setText(Path(filename).name)

    def closeEvent(self, event):
        worker = getattr(self, "worker", None)
        if worker is not None and worker.isRunning():
            worker.cancel()
            event.ignore()
            from PySide6.QtCore import QTimer
            QTimer.singleShot(100, self.close)
            return
        event.accept()

    def _on_finished(self, ok, msg):
        if ok:
            self.stack.setCurrentIndex(2)
            self.btn_cancel.hide()
            self.btn_action.setText("Finish")
            self.btn_action.setEnabled(True)
        else:
            self.prog_status.setText(f"Error: {msg}")
            self.btn_cancel.setEnabled(True)

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = InstallerWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
