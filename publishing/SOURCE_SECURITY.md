# FINDEX security policy

This is a source-code publication of FINDEX 3.0.1, not a security-certified binary
release. No fixed response-time or long-term support guarantee is offered.

## Reporting

Use this repository's Security → Advisories → Report a vulnerability option when
private vulnerability reporting is enabled. The maintainer must enable that
GitHub setting; adding this file does not enable it. If the private route is not
available, do not disclose exploit details or evidence in a public issue. A real
private contact must be established first; no email address is invented here.

Provide the version, affected module, impact and a minimal reproduction using
synthetic files. Never include private case files, credentials or unredacted paths.

## Known native findings — reviewed 2026-11-14

The source pins PySide6 6.12.0. Qt's native advisories include affected components
at that version. Four reviewed triggering modules are absent from the local
portable artifact; Qt Core and Qt Network are present. The affected Core CBOR
and Network HTTP APIs have no direct application-source use identified, but
transitive native reachability has not been proved absent. They remain open risks.

PySide6 6.12.1 is not available on PyPI at the review date. A 6.11.3 wheel exists,
but the cited Network advisory does not identify 6.11.3 as a fix, so downgrading
does not establish a complete resolution. Dependencies were not silently changed.
See NATIVE_DISTRIBUTION_REVIEW.md for the component inventory and dispositions.

## Handling evidence

FINDEX is intended to read evidence and write separate reports. Use a forensic
write blocker/read-only snapshot where required: application read-only opens
cannot guarantee Windows access timestamps or all parser behavior. A matching
hash does not establish authenticity, ownership or provenance. MD5/SHA-1 are
retained for legacy compatibility; use an appropriate modern reference hash.

Treat native metadata/image/media parsers as an attack surface. This is not a
sandbox for hostile files. Reports may intentionally contain paths and metadata;
review before sharing. Existing outputs require consent before replacement.

The local historical Defender scan and passing tests are evidence of those
specific checks only, not proof of malware absence or forensic reliability.
Noncommercial distribution still has third-party license conditions. This
repository contains no bundled native executable/runtime distribution.
