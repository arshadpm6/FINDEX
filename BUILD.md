# Building FINDEX 3.0.1

Use Windows x64, Python 3.12 x64 and an isolated environment. The exact reviewed
dependency closure is `requirements-release.lock`; runtime dependencies are in
`requirements.txt`, and packaging dependencies in `requirements-build.txt`.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --require-hashes -r requirements-release.lock
.\.venv\Scripts\python.exe scripts\check_build_environment.py
.\.venv\Scripts\python.exe scripts\build_release.py --output-dir release/FINDEX-v3.0.1
```

The output directory must be new or empty. Existing release folders and the old
root installer are preserved. Work is under `build`, `dist` and `scratch`. The
build does not install dependencies implicitly or change Windows security policies.
Use the Python command when PowerShell scripts are restricted.

## Source checkout native tools

The public source archive excludes native helper binaries. Provision these before
building; the application does not automatically download them:

- `external/exiftool/exiftool.exe` and the complete matching `exiftool_files`,
  including `perl.exe`, `exiftool.pl`, `lib` and all notices. The local reviewed
  bundle reports ExifTool 13.59. Obtain Windows packages from <https://exiftool.org/>.
- `external/ffmpeg/ffprobe.exe`: the local reviewed Gyan 9.0.2 essentials build.
  Archive URL/checksum and upstream commit are in `external/ffmpeg/SOURCE.json`.
  Obtain it from <https://www.gyan.dev/ffmpeg/builds/> and verify that checksum.
  If the rolling release changes, do not silently accept a different archive.

The working checkout already has these. Generated `NATIVE_COMPONENTS.json` records
exact native file hashes for comparison, not authenticated upstream signatures.
A clean source clone requires manual native-tool provisioning; release tests
require matching tools. Keep licenses and corresponding source with distributions.

## Validation and outputs

The build validates dependency versions, collects notices, runs worker/UI/export,
theme/branding/startup tests, and builds a windowed portable app with UPX disabled.
It checks Windows version/icon/asInvoker resources, startup and bundled helpers,
verifies the payload, tests installation/uninstallation in an owned Unicode scratch
directory, then builds Setup. Tests do not create real shortcuts or registry entries.
Outputs include portable/source ZIPs, documentation, manifest and SHA-256 checksums.

```powershell
.\scripts\verify_release.ps1 -Manifest .\release\FINDEX-v3.0.1\RELEASE_MANIFEST.json
```

Network is required to obtain build dependencies, not to run FINDEX. Input versions
are pinned; bit-for-bit identical PyInstaller output is not claimed. No signing
key is supplied. Complete [PUBLIC_RELEASE.md](PUBLIC_RELEASE.md) before publishing.

The frozen application has a fixed synthetic Excel diagnostic for build verification:
`FINDEX.exe --check-excel <new-empty-output-directory>`. The directory must not
already exist. It creates only synthetic Excel/HTML reports and a result.json;
it does not read evidence, change preferences or open the main window. Portable
smoke tests require both clean and simulated legacy-NumPy cases to export and
reopen real workbooks successfully.

The Windows specs exclude foreign root-level ICU libraries and use the operating
system's ICU API. The build puts Windows system paths first so unrelated tools
such as Poppler cannot inject an incompatible ICU copy. No system DLL is replaced.
