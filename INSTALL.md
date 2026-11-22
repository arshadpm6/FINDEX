# Installing FINDEX 3.0.1

## Windows installer

Download `FINDEX_v3.0.1_Setup.exe` from
https://github.com/arshadpm6/FINDEX/releases/tag/v3.0.1-core.
This preview core package includes Python and the desktop runtime and targets Windows x64.
It does not bundle ExifTool or ffprobe.

Verify the EXE against the release's SHA256SUMS.txt, close any running FINDEX,
and run Setup. Choose a dedicated FINDEX application folder. Desktop/Start Menu
shortcuts and per-user Installed Apps registration are optional. No administrator
elevation or internet connection is needed for the core installation.

Launch FINDEX.exe or the installed shortcut. Hashing, inventory, Excel and HTML
reports work immediately. Metadata Explorer requires a separately configured
ExifTool; optional media-duration extraction requires ffprobe. See the detailed
[README](README.md#optional-metadata-and-media-tools).

For portable use, extract the complete portable ZIP and retain the _internal folder.
To update, close the application and run Setup against the existing installation.
Keep install_manifest.json. To uninstall, use Installed Apps if registered, or
uninstall.bat in the installation folder. Retain reports/settings you need before updating.

The package is unsigned. Investigate actual security detections; do not disable
Defender/SmartScreen or organization controls. Checksums detect changed bytes but
are not an authenticated signature.

## Source checkout

Use Python 3.12 x64 on Windows:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe "FINDEX.py"
```

Configure official ExifTool in Settings for metadata extraction. Retain its
complete matching Windows runtime. For durations, install ffprobe on PATH before
launch, or provision external/ffmpeg/ffprobe.exe. Native helpers are not committed.
See BUILD.md for packaging and CORE_RELEASE.md for this release's component scope.
