# Running FINDEX from source

This GitHub publication does not include native executables, Qt DLLs or an installer.
Use Windows x64 and Python 3.12 x64.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe "File Inventory Pro.py"
```

The legacy source filename is retained for build compatibility; the application
identity and executable name are FINDEX.

For Metadata Explorer, obtain ExifTool from https://exiftool.org/ and configure
its path in Settings, or provision the matching complete Windows ExifTool bundle
under `external/exiftool` as described in BUILD.md. For media duration, provision
ffprobe under `external/ffmpeg/ffprobe.exe` or on PATH. Inventory and hash features
remain available when optional media-duration extraction is unavailable.

The previously built local standalone package bundles those tools, but it is not
part of this source publication. Do not redistribute it until the native binary
requirements in NATIVE_DISTRIBUTION_REVIEW.md are resolved.

Current PySide6 pins are disclosed in SECURITY.md. Use controlled/synthetic files
for evaluation; no hostile-file sandbox or full security validation is claimed.

For local packaging, follow BUILD.md, retain third-party notices and review the
separate source/license conditions before distributing any resulting executable.
