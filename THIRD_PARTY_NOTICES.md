# FINDEX third-party components

FINDEX is distributed separately from the third-party components below. Their
copyrights and license conditions continue to apply. No ownership of them is claimed.

| Component | Use | License/source information |
|---|---|---|
| Python 3.12 | Bundled interpreter | PSF license; https://www.python.org/downloads/source/ |
| PySide6 / Shiboken / Qt 6.12.0 | Desktop interface and file previews | LGPL/GPL/commercial alternatives and component-specific notices; https://doc.qt.io/qtforpython-6/licenses.html and https://download.qt.io/official_releases/QtForPython/pyside6/ |
| openpyxl 3.1.5 | Excel reports | MIT; https://openpyxl.readthedocs.io/ |
| Pillow 13.0.0 | Image handling | HPND and bundled library notices; https://python-pillow.github.io/ |
| ExifTool 13.59 and Windows Perl runtime | Metadata extraction | ExifTool uses the same terms as Perl; retain the GPL text, `Licenses_Strawberry_Perl.zip`, library notices and ExifTool source included under `external/exiftool/exiftool_files`; https://exiftool.org/ |
| ffprobe 9.0.2 Gyan build | Media durations | This bundled build enables GPL components. Retain `external/ffmpeg/LICENSE` and `SOURCE.json`; https://www.gyan.dev/ffmpeg/builds/ and https://www.ffmpeg.org/legal.html |
| PyInstaller 6.11.1 | Packaging/bootloader | GPL with bootloader exception; https://pyinstaller.org/en/stable/license.html |

Release preparation must include the applicable license texts and corresponding source
for the actual binaries distributed, including applicable FFmpeg build dependencies.
A URL to an upstream source tree alone does not establish full compliance. The current
FFmpeg SOURCE.json records its upstream commit and download provenance; it is not a
complete corresponding-source distribution. Arrange that source distribution before
publishing the installer. Likewise review the applicable Qt module conditions and
replacement/relinking rights. FINDEX keeps Qt libraries in the separate `_internal`
folder, without obfuscation or an integrity check that blocks user replacement.

The owner selected the MIT license for original FINDEX code; see LICENSE.
It does not replace third-party licenses or certify legal compliance.

The reviewed PySide6 wheel metadata lists LGPL/GPL alternatives, while its shipped
license directory includes a commercial-license notice. FINDEX does not claim a
commercial Qt license. GNU LGPLv3/GPLv2/GPLv3 texts from the SPDX license list are kept
under `resources/third_party/Qt-license-texts`, with download URLs and checksums.
These texts supplement the wheel notices; they do not complete Qt corresponding-
source/module-specific obligations or substitute for reviewing the actual binaries.
