# FINDEX — File Intelligence & Examination

**Know Every File.** · Version 3.0.1 · Windows desktop application

FINDEX is my personal software project for file inventory, hash verification and
metadata examination. I am sharing its source code for learning, feedback and
collaboration. Original FINDEX code is licensed under the [MIT License](LICENSE).
The 2026-11-22 source correction restores Excel export with incompatible optional
NumPy installs and adds Settings-controlled Automatic/KB/MB inventory report sizes.
MIT also permits commercial reuse; describing this as a personal/noncommercial
project does not add a noncommercial restriction to that license.

## Modules

- **Hash Verification:** calculate hashes, compare expected values, compare two
  files and export verification reports. MD5/SHA-1 remain for legacy workflows;
  use SHA-256 or another suitable modern algorithm for new reference hashes.
- **File Inventory:** scan folders, choose report columns, filter file categories
  and generate Excel/HTML inventories.
- **Metadata Explorer:** examine metadata using ExifTool, with optional hashes,
  integrity checks and HTML, TXT or JSON exports.

FINDEX uses Python/PySide6. It operates locally without an application login,
cloud account or telemetry. The desktop theme, navigation and existing three
modules are preserved in this source publication.

## Getting started

See [INSTALL.md](INSTALL.md) for running from source and [BUILD.md](BUILD.md) for
local packaging. Python dependencies are pinned in `requirements.txt`; the
reviewed Windows build closure is in `requirements-release.lock`.
ExifTool must be provisioned separately for this source checkout; ffprobe is used
for media duration. Neither native tool nor the Qt runtime is committed here.

**This publication contains source code, not a public installer release.** Do not
attach the earlier local installer/portable package until their native component
source/license requirements and release checks are complete. See
[NATIVE_DISTRIBUTION_REVIEW.md](NATIVE_DISTRIBUTION_REVIEW.md).

## Status and limits

The 2026-11-22 correction is covered by worker/UI regressions, real workbook
round-trips and persisted report-size settings tests. Build tests also verify
the frozen EXE's Excel support with a simulated incompatible legacy NumPy module.
See FINAL_RELEASE_VERIFICATION.md for the recorded completed checks and skips.
These are not a guarantee for every PC or a certification of forensic reliability
or malware absence. The native review is historical, not a claim that the latest
build received a new native advisory clearance.

The current Qt 6.12.0 dependency has published native advisories. The component
and application-source review records what is absent, what is present, and what
remains unresolved. Do not describe the project as fully patched or security-certified.

Use write blockers/read-only snapshots when evidence preservation requires them.
Matching hashes do not prove authenticity or provenance. Reports can contain
sensitive paths and metadata. Use synthetic files when reporting bugs or sharing
screenshots; do not upload case evidence or personal data.

## Contributing and reporting

Read [CONTRIBUTING.md](CONTRIBUTING.md), [SECURITY.md](SECURITY.md),
[CHANGELOG.md](CHANGELOG.md) and [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
Public source publishing steps are in [PUBLIC_RELEASE.md](PUBLIC_RELEASE.md).
Third-party licenses remain separate from FINDEX's MIT license.
