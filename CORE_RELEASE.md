# FINDEX 3.0.1 core Windows release

This release adds downloadable Setup and portable assets without redistributing
the previously assembled ExifTool/Perl or GPL-enabled ffprobe bundles.

## Included and excluded

Included: original FINDEX application, Python 3.12 runtime, PySide6/Shiboken/Qt
6.12.0 core GUI libraries, openpyxl, Pillow, required runtime DLLs, project assets,
and component license notices.

Excluded: ExifTool/Perl, FFmpeg/ffprobe, Qt Network, QML/Quick, PDF,
Virtual Keyboard and unrelated Qt plugins. Hashing, inventories and report
exports work without helper executables. Metadata requires user-provisioned
ExifTool; media duration requires user-provisioned ffprobe.

The original full-tool build scripts and historical NATIVE_DISTRIBUTION_REVIEW.md
remain available. Their older binary review does not describe this core artifact.

## Third-party source and replacement

The release offers official, unmodified source archives for Qt Base, Qt SVG,
Qt Image Formats and PySide/Shiboken 6.12.0 alongside the binary assets.
THIRD_PARTY_SOURCES.json records each upstream URL, filename and SHA-256.
These archives include their upstream build instructions and third-party notices.
The source inputs correspond to the published 6.12.0 runtime releases; this
project has not patched those libraries.

Qt/PySide use their applicable LGPL/GPL terms; included license texts remain in
resources/third_party and in the installed _internal/resources/third_party tree.
The original FINDEX source stays MIT. You may modify/rebuild the covered
libraries and replace the compatible dynamic DLLs/extensions in _internal;
the application does not enforce library signatures or prevent replacement.
Use matching Windows x64/CPython 3.12-compatible bindings and ABI versions.
See the upstream CMake/Qt and Qt-for-Python build documentation in the supplied
archives and https://doc.qt.io/qtforpython-6/building_from_source/windows.html.

The one-file Setup temporarily extracts its own GUI runtime. To use rebuilt
libraries for installation, run installer_wizard.py from the supplied FINDEX
source with a compatible Python/PySide environment after regenerating
dist/app_payload.zip and its manifest using scripts/package_payload.py against the
extracted portable folder, or rebuild FINDEX_Core_Setup.spec. For the installed or
portable application, DLLs remain separate and replaceable. Source and build specs
are provided; no restrictions on debugging/reverse engineering library modifications
are added by FINDEX.

Python and permissively licensed dependencies retain their notices. The official
source URLs are recorded for auditing and rebuilding. No legal certification is asserted.

## Known limits

Qt's current public PySide6 release was checked as 6.12.0 on 2026-11-22.
The Qt Core CBOR-to-JSON issue CVE-2026-105924 remains applicable to that library;
no direct application use of the triggering API was found. Qt Network is absent
from this core binary. See https://wiki.qt.io/List_of_known_vulnerabilities_in_Qt_products.
Use controlled files for evaluation: FINDEX is not a hostile-file sandbox.

The exact release verification record is supplied as VERIFICATION.md.
Local/offscreen tests do not substitute for a disposable clean-machine test,
a current endpoint-protection scan, or code signing. The release is marked preview.
