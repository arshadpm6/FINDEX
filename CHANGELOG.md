# Changelog

## 3.0.1 — corrections, 2026-11-22

- Keep Excel export working when an incompatible optional NumPy installation is
  present; retry openpyxl without NumPy instead of disabling Excel functionality.
- Add a persistent File Inventory size setting: Automatic, KB or MB, applied to
  Excel/HTML file values, column headings and report totals; raw bytes are retained.
- Make empty-column Excel exports use the same minimal columns as HTML exports.
- Verify real workbook exports in the frozen EXE, both with a clean bundle and
  a simulated incomplete legacy NumPy installation.
- Prevent an unrelated Poppler ICU DLL on the build PATH from shadowing the
  Windows ICU API required by Qt in both app and installer bundles.
- Consolidate the latest local executable/installer and GitHub source; remove
  obsolete build copies, temporary environments, backups and generated reports.

## 3.0.1 — release preparation, 2026-11-14

Reviewed local release candidate; no GitHub publication or earlier release history
is inferred.

- FINDEX identity, desktop icons, About information and HTML report branding.
- Preserved Hash Verification, File Inventory and Metadata Explorer modules.
- Corrected hash/result handling, worker errors/cancellation and source-change checks.
- Escaped report content and protected spreadsheet/CSV cells against formula input.
- Protected evidence paths and staged approved report replacements atomically.
- Replacement permission only for actual existing output files, not existing folders.
- HTML/TXT metadata exports no longer create JSON sidecars; JSON/raw JSON are opt-in.
- A4 metadata print layout, theme-aware dialogs and consistent loading animation.
- Preserved hidden subprocess/startup behavior; no automatic folder opening.
- Payload verification, manifest-scoped uninstall, pinned build inputs and checksums.
- MIT license, installation/build instructions and public release checklist.
- Corrected collection of British-spelled LICENCE files in bundled third-party notices.
