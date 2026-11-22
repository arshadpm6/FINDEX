# FINDEX — File Intelligence & Examination

**Know Every File. · Version 3.0.1 · Windows x64**

FINDEX is a desktop application for file inventories, checksum verification, and
metadata examination. It processes files locally and exports reports you can keep
or share. Original FINDEX code is available under the [MIT License](LICENSE).

## Download FINDEX for Windows

**[Download the Windows installer](https://github.com/arshadpm6/FINDEX/releases/download/v3.0.1-core/FINDEX_v3.0.1_Setup.exe)**

Open the [release page](https://github.com/arshadpm6/FINDEX/releases/tag/v3.0.1-core)
for the installer, SHA-256 checksums, portable ZIP, verification results, and
third-party source materials. Under **Assets**, choose `FINDEX_v3.0.1_Setup.exe`.
GitHub's automatically generated **Source code** archives contain developer source,
not the installed Windows application.

This is a **preview core package**. Python and the desktop runtime are included;
you do not need to install Python. Hash verification, inventories, and Excel/HTML
report generation work without additional tools. **Metadata Explorer requires a
separate ExifTool installation. Media durations require optional ffprobe.**
These native helper executables are not bundled in this download. See their
[setup instructions](#optional-metadata-and-media-tools) below.

While this repository is private, GitHub requires an account with repository
access to download its files. Making the repository public enables public
downloads; an installer upload does not change repository visibility.

## What FINDEX does

| Module | Main tasks | Reports |
| --- | --- | --- |
| Hash Verification | Compute checksums, compare against expected values, compare two files | Excel, HTML, CSV, checksum files |
| File Inventory | Recursively list files; select columns and categories; include sizes, dates, paths, and optional durations | Excel and HTML |
| Metadata Explorer | Read metadata through ExifTool, inspect categories and raw extraction output, optionally calculate integrity hashes | HTML, TXT, JSON |

- Supported checksum choices include MD5, SHA-1, SHA-256, and SHA-512.
  Use SHA-256 or another suitable modern algorithm for new reference checksums.
  MD5 and SHA-1 remain available for older workflows.
- Settings control output locations, hash choices, and inventory size display.
  Inventory reports support **Automatic**, **KB**, and **MB**; save the setting
  before generating a new report. KB and MB use powers of 1024.
- The interface includes dark/light themes, progress feedback, cancellation,
  and report-opening actions.
- Existing report files require replacement confirmation. Keep generated reports
  outside folders containing your source files.

## Install, launch, update, and uninstall

1. Download the Setup EXE and `SHA256SUMS.txt` from the same release.
2. Optionally verify the checksum using the command below.
3. Run `FINDEX_v3.0.1_Setup.exe` and select a dedicated folder named **FINDEX**.
   The default is a per-user application folder.
4. Choose whether to create Desktop/Start Menu shortcuts and an Installed Apps
   entry, then finish installation.
5. Launch **FINDEX** from a shortcut or run `FINDEX.exe` in the installation folder.

The target is **64-bit Windows**. The installer does not require a system Python
installation. It installs per user and does not request administrator elevation.
Installation and core features work offline. Downloading optional tools requires
internet access; FINDEX does not download them automatically.

For an update, close FINDEX and run the new installer against the existing
application folder. Keep `install_manifest.json`: it records which files belong
to the installation. An older folder named **File Inventory Pro** is also
recognized. Back up reports and settings before updating.

To uninstall, use **Windows Settings → Apps → Installed apps → FINDEX** if that
entry was selected, or run `uninstall.bat` in the application folder. Uninstall
uses the installation manifest to remove application files while retaining
unrelated files. Managed Windows policy can prevent the uninstall script; do not
disable organization security controls to work around that.

### Portable use

Extract `FINDEX_v3.0.1_Portable.zip` into a new folder and run
`FINDEX/FINDEX.exe`. Keep the accompanying `_internal` folder intact.
The portable package has the same optional-tool requirements as Setup.
Close FINDEX before replacing a portable folder during an update.

### Verify a download

In PowerShell, from the download folder:

```powershell
Get-FileHash -LiteralPath .\FINDEX_v3.0.1_Setup.exe -Algorithm SHA256
```

Compare the complete value with the matching filename in the release's
`SHA256SUMS.txt`. A checksum detects changed bytes against a trusted reference;
it does not establish the publisher's identity. This release is unsigned.
If Windows or antivirus software reports a detection, retain the detection details
and investigate it. Do not disable protection or add exclusions to run FINDEX.

## Optional metadata and media tools

### ExifTool: required only for Metadata Explorer

1. Obtain the official Windows package from [exiftool.org](https://exiftool.org/).
2. Extract the **complete package**, including its matching `exiftool_files`
   directory. Follow the provider's executable-renaming instructions if its
   launcher is named `exiftool(-k).exe`.
3. Open FINDEX **Settings**, configure the ExifTool executable path, and use the
   configuration dialog's version/status check.
4. Select a harmless sample file in Metadata Explorer and confirm extraction.

Keep the launcher and matching runtime together. Alternatively, FINDEX can find
an ExifTool installation on PATH or in its supported local tool locations.
The previous source tests used ExifTool 13.59; this core installer does not
redistribute that Windows/Perl bundle.

### ffprobe: optional media durations

Obtain an appropriate Windows ffprobe distribution through the links on the
[FFmpeg download page](https://ffmpeg.org/download.html). Install it on PATH
before starting FINDEX, or provision `external/ffmpeg/ffprobe.exe` under the
application's `_internal` directory. Restart FINDEX after changing PATH.
Follow that provider's license and source instructions.

Without ffprobe, hashing and inventories still work; duration values are omitted
and an unavailable-tool warning can appear when duration extraction is requested.
The previously assembled GPL-enabled Gyan binary is deliberately not included.

## A quick first run

Use synthetic or non-sensitive files while learning the interface.

- **Verify a file:** open Hash Verification, add a file, choose SHA-256, and compute
  its checksum. Supply an expected checksum when you want a comparison.
- **Create an inventory:** select a folder, choose columns/categories and an output
  location, select Excel and/or HTML, then start the scan.
- **Inspect metadata:** configure ExifTool first, open Metadata Explorer, select a
  sample file, review the extracted categories, and select report formats.
  Raw ExifTool JSON is a separate optional output, not required for HTML/TXT reports.

Large folders and slow/network storage take longer. Avoid changing files during
a scan. Review warnings and skipped/error counts before relying on a report.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| ExifTool is missing | Configure the executable path; retain the complete matching runtime; run the version/status check |
| Media durations are empty | Install ffprobe on PATH or in the supported local path, then restart FINDEX |
| Excel export fails | Check the output path and permissions; close an existing workbook; include the exact error when reporting a bug |
| Setup rejects a folder | Choose a dedicated FINDEX folder; do not install into a general documents/downloads directory |
| Portable application does not start | Fully extract the ZIP and retain `_internal`; confirm Windows x64 and review security-product details |
| A report already exists | Choose a new filename or approve replacement only if you intend to overwrite that report |
| Download is unavailable | Check repository access/visibility and open the release's Assets list |

Do not upload private evidence, credentials, report contents, or full personal
paths in an issue. Describe the FINDEX version, Windows version, exact message,
and a reproduction using harmless sample files.

## Run from source

Use Python 3.12 x64 on Windows:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe "FINDEX.py"
```

Runtime
dependencies are PySide6, openpyxl, and Pillow. Native helper executables are
excluded from source control and must be provisioned separately when needed.
See [INSTALL.md](INSTALL.md) and [BUILD.md](BUILD.md).

For the core installer, use [scripts/build_core_release.py](scripts/build_core_release.py)
with the hash-pinned build environment. The release includes a file manifest,
dependency source materials, and recorded checks. Existing full-tool packaging
scripts are retained for local use; they are not the distribution recipe for this
core release.

## Security, privacy, and verification limits

FINDEX runs locally without an application login, telemetry, or a cloud service.
Reports can contain sensitive paths and metadata; review them before sharing.
Use a write blocker/read-only snapshot when evidence-preservation rules require it.
Read-only application operations cannot guarantee filesystem access timestamps or
the safety of every native parser. Matching hashes do not prove provenance.

The core package uses PySide6/Qt 6.12.0. A known Qt Core CBOR-to-JSON advisory remains
disclosed; the application does not directly call that API. Unused Qt Network,
QML, PDF, and Virtual Keyboard components are excluded from this package.
No claim of complete vulnerability freedom, malware certification, code signing,
or forensic certification is made. The verification report distinguishes local
fixture checks from clean-machine or endpoint-protection checks.

Read [SECURITY.md](SECURITY.md), [CORE_RELEASE.md](CORE_RELEASE.md), and the release's
`VERIFICATION.md` before distribution or evaluation with sensitive material.

## Project documentation and contributing

- [Installation and tool setup](INSTALL.md)
- [Build instructions](BUILD.md)
- [Core package and third-party source details](CORE_RELEASE.md)
- [Changelog](CHANGELOG.md) and [release notes](RELEASE_NOTES.md)
- [Contributing](CONTRIBUTING.md) and [security](SECURITY.md)
- [Third-party notices](THIRD_PARTY_NOTICES.md)

Original FINDEX source is MIT licensed. Third-party runtimes retain their own
licenses. Qt/PySide source archives, notices, and replacement/rebuild instructions
are provided with the core release; FINDEX's MIT license does not replace those terms.
