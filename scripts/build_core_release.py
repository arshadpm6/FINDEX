"""Build the Windows core preview without bundling ExifTool/Perl or FFmpeg."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import time
import zipfile

ROOT = Path(__file__).resolve().parents[1]

def sha256(path):
    with path.open("rb") as stream:
        result = hashlib.file_digest(stream, "sha256").hexdigest()
    return result

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", default="release/FINDEX-v3.0.1-core")
    args = parser.parse_args()
    release = (ROOT / args.output_dir).resolve()
    if not release.is_relative_to(ROOT) or release == ROOT:
        raise SystemExit("Use an output directory inside the checkout")
    if release.exists() and any(release.iterdir()):
        raise SystemExit("Existing release preserved: use a new or empty directory")
    release.mkdir(parents=True, exist_ok=True)
    (ROOT / "scratch").mkdir(exist_ok=True)
    environment = os.environ.copy()
    environment.pop("PYTHONPATH", None)
    environment["PYTHONNOUSERSITE"] = "1"
    windows = Path(environment.get("SystemRoot", r"C:\Windows"))
    environment["PATH"] = os.pathsep.join((str(windows / "System32"), str(windows), environment.get("PATH", "")))
    results = []

    def run(*arguments):
        started = time.monotonic()
        print("RUN:", " ".join(arguments), flush=True)
        subprocess.run([sys.executable, *arguments], cwd=ROOT, env=environment, check=True)
        results.append({"command": list(arguments), "result": "passed", "seconds": round(time.monotonic() - started, 2)})

    run("scripts/check_build_environment.py")
    run("scripts/collect_notices.py")
    for name in ("test_security_hardening.py", "test_gui_corrections.py",
                 "test_message_box_theme.py", "test_report_branding.py",
                 "test_export_selection.py", "test_inventory_sizes.py", "check_startup_windows.py"):
        run("tests/" + name)
    run("-m", "PyInstaller", "--clean", "--noconfirm", "FINDEX_Core.spec")
    application = ROOT / "dist/FINDEX"
    internal = application / "_internal"
    assert not (internal / "external").exists(), "Optional helper bundles must be absent"
    forbidden = ("Qt6Network.dll", "Qt6Pdf.dll", "Qt6VirtualKeyboard.dll", "Qt6Qml.dll", "Qt6Quick.dll")
    assert not any((internal / "PySide6" / name).exists() for name in forbidden)
    run("scripts/check_binary_identity.py", str(application / "FINDEX.exe"))
    run("tests/smoke_portable.py", "--core")
    run("scripts/package_payload.py")
    run("tests/test_payload_install.py", "--core")
    run("-m", "PyInstaller", "--clean", "--noconfirm", "--distpath", "scratch/setup-stage", "FINDEX_Core_Setup.spec")
    setup = ROOT / "scratch/setup-stage/FINDEX_v3.0.1_Setup.exe"
    run("scripts/check_binary_identity.py", str(setup))
    run("tests/smoke_installer.py", str(setup))
    shutil.copy2(setup, release / setup.name)
    docs = ("README.md", "INSTALL.md", "LICENSE", "CORE_RELEASE.md", "THIRD_PARTY_NOTICES.md", "SECURITY.md", "RELEASE_NOTES.md")
    for name in docs:
        shutil.copy2(ROOT / name, release / name)
    payload = json.loads((ROOT / "dist/app_payload.manifest.json").read_text(encoding="utf-8"))
    (release / "APPLICATION_PAYLOAD.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    with zipfile.ZipFile(release / "FINDEX_v3.0.1_Portable.zip", "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(application.rglob("*")):
            if path.is_file():
                archive.write(path, "FINDEX/" + path.relative_to(application).as_posix())
    original = json.loads((ROOT / "SOURCE_FILES.json").read_text(encoding="utf-8"))
    selected = {entry["path"] for entry in original}
    selected.update(("FINDEX_Core.spec", "FINDEX_Core_Setup.spec", "CORE_RELEASE.md", "scripts/release_components.py", "scripts/build_core_release.py"))
    records = [{"path": name, "size": (ROOT / name).stat().st_size, "sha256": sha256(ROOT / name)} for name in sorted(selected)]
    (ROOT / "SOURCE_FILES.json").write_text(json.dumps(records, indent=2), encoding="utf-8")
    with zipfile.ZipFile(release / "FINDEX_v3.0.1_Source.zip", "w", zipfile.ZIP_DEFLATED) as archive:
        for entry in records:
            assert Path(entry["path"]).suffix.lower() not in (".exe", ".dll", ".pyd", ".key", ".pem", ".pfx")
            archive.write(ROOT / entry["path"], entry["path"])
        archive.write(ROOT / "SOURCE_FILES.json", "SOURCE_FILES.json")
    report = {"platform": platform.platform(), "python": sys.version, "checks": results,
              "source_sha256": sha256(ROOT / "FINDEX.py"),
              "setup_sha256": sha256(release / setup.name),
              "optional_native_helpers": "not bundled",
              "unused_qt_components": "Network/PDF/QML/Quick/VirtualKeyboard omitted",
              "not_performed": ["disposable clean Windows VM", "endpoint-protection scan", "Authenticode signing"]}
    (release / "VERIFICATION.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    text = "# FINDEX 3.0.1 core preview verification\n\n"
    text += "Built on 2026-11-22 using " + platform.platform() + ".\n\n"
    text += "The following local synthetic/offscreen checks passed:\n\n"
    text += "".join("- " + " ".join(entry["command"]) + "\n" for entry in results)
    text += "\nVerified: optional native helpers and unused Qt modules are absent; Setup and app identity/resources; real frozen Excel/HTML/hash report exports including incompatible optional NumPy; payload installation/uninstallation fixtures and preservation checks; Setup startup.\n"
    text += "\nNot performed: a disposable clean Windows VM trial, a current endpoint-protection scan, or Authenticode signing. Tests use controlled local fixtures. Startup checks are offscreen. This is a preview, not a safety or forensic certification.\n"
    text += "\nKnown Qt Core CBOR advisory remains disclosed in CORE_RELEASE.md.\n"
    (release / "VERIFICATION.md").write_text(text, encoding="utf-8")
    print("CORE RELEASE READY:", release, flush=True)

if __name__ == "__main__":
    main()
