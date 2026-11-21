# Contributing to FINDEX

Please describe the affected module and expected/observed behavior before making
large changes. Preserve Hash Verification, File Inventory and Metadata Explorer,
existing branding/themes and the hidden subprocess/startup behavior.

Use synthetic test files. Do not submit case evidence, private reports, credentials,
personal paths or binary dependencies. Changes must not write to source evidence.

Use Python 3.12 on Windows and the isolated environment described in BUILD.md.
Full regression tests require the native tools provisioned locally; they are not
committed in this source publication. Run tests appropriate to the change and
record skips/limitations honestly. Never mark a missing-dependency test as passed.

Original project contributions are made under this repository's MIT license unless
separately agreed. Keep third-party attribution and licenses. Report security
issues through SECURITY.md, not public exploit reports.
