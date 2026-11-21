# Native distribution review — FINDEX 3.0.1

Reviewed 2026-11-14. Application logic, UI, branding and native dependencies were not changed.

## Decision

Prepare the original FINDEX source for GitHub, without native binaries. Keep the
existing Setup EXE and portable ZIP out of the public commit/downloads. Personal
or noncommercial intent does not remove binary redistribution conditions.
This review does not certify legal compliance or absence of vulnerabilities.

## Qt advisory review

The exact portable archive has SHA-256 `810867067b5c64d8b0a326cdde7a593db2245ba819a1ab08550a0078836c0c1e`. Its native Qt inventory and
application-source hashes are in NATIVE_DISTRIBUTION_REVIEW.json. The reviewed
advisory findings relevant to Qt 6.12.0 are:

| Advisory | Triggering component | Local artifact disposition |
|---|---|---|
| CVE-2026-116337 | Qt Network HTTP responses | Present; no direct trigger found; remains open |
| CVE-2026-110944 | Qt HTTP Server | Triggering module absent from artifact |
| CVE-2026-114366 | Qt WebSockets | Triggering module absent from artifact |
| CVE-2026-111487 | Qt3D assimp importer | Triggering module absent from artifact |
| CVE-2026-100140 | Qt5Compat XML | Triggering module absent from artifact |
| CVE-2026-105924 | Qt Core CBOR conversion | Present; no direct trigger found; remains open |

Source analysis found no direct application references to the checked triggering
APIs. That does not prove all native/transitive paths unreachable. Core and
Network findings remain open; their presence is not cleared by a pip-audit result.
The four absent-module dispositions apply to this exact artifact, not every Qt installation.

PyPI metadata currently lists PySide6 6.12.0 as latest.
6.12.1 is unavailable. 6.11.3 exists, but the cited HTTP-response advisory does
not list that branch as fixed; a downgrade does not establish complete resolution.
Keep the tested pin, disclose the findings, and re-evaluate compatible patches.
Reference: https://wiki.qt.io/List_of_known_vulnerabilities_in_Qt_products

## FFmpeg distribution review

FINDEX uses ffprobe as a separate executable for media duration. The current
Gyan 9.0.2 build is static and enables GPL/version3 components. The binary hash,
configuration, upstream revision, archive provenance and 37 recorded
external-library version lines are retained in the JSON report.

The provider's 9.0.2 GitHub release assets are binary archives; the checked asset
list contains no complete corresponding-source bundle. An upstream FFmpeg commit
link alone does not account for all statically included dependencies, patches and
build materials. These have not been collected/verified. Do not label the existing
installer compliant merely by adding a license file or a generic written offer.
See THIRD_PARTY_NOTICES.md and https://www.ffmpeg.org/legal.html.

The prepared GitHub tree omits ffprobe, ExifTool/Perl binaries and Qt runtime
libraries. This avoids publishing those binary copies in the source commit; it
does not clear the separately built installer for distribution. Users provision
their local dependencies through the original providers. FINDEX's original source
keeps MIT; third-party components retain their terms.

## Next steps for binary publication

Obtain the exact native corresponding-source/build materials, verify coverage,
and make them available with the binary distribution as applicable. Reassess Qt
patches/applicability, test the resulting build and repeat exact-artifact scans.
A maintainer request draft is in publishing/FFMPEG_SOURCE_REQUEST.md; it has not
been sent. No original source evidence was used or modified in these checks.
