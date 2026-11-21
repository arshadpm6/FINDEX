# FINDEX 3.0.1 — corrections and local verification

Correction date: 2026-11-22. Local executable/installer; GitHub publication remains source-only.

## Identified issues and corrections

- Excel: openpyxl was installed, but an incomplete optional NumPy module could raise
  an AttributeError during import. The previous guard disabled Excel support and
  displayed an openpyxl-required error. The loader now retries openpyxl without
  incompatible optional NumPy, removes only newly loaded partial openpyxl modules,
  and restores the original NumPy entry. Missing openpyxl/unrelated errors remain visible.
- Settings now has a separate File Inventory report size selector: Automatic, KB or
  MB. Save Settings to apply it to Excel/HTML file cells, forced-unit headings and
  totals. Raw bytes, hashes, filters and evidence handling are retained. Units use
  1024 bytes per KB and 1024 KB per MB; forced units show two decimals. The existing
  KB/MB hash-stream buffer setting is unchanged.
- Empty-column Excel reports use the same minimal serial/name columns as HTML.
- The packaged-app check identified a foreign Poppler ICU DLL pulled from the
  build PATH. It lacked the Windows ICU exports Qt imports, so QtWidgets could
  not load in that bundle. App/installer specs now exclude those foreign ICU
  files and the build prioritizes Windows system paths. The fixed temporary
  EXE passed its synthetic report check after removal; the final bundle repeats it.
- Build checks now generate/reopen real Excel workbooks and HTML in the frozen EXE,
  including a simulated incompatible legacy NumPy module. The fixed --check-excel
  diagnostic writes synthetic reports only into a new directory and opens no GUI.

## Validation

Source security/worker, GUI and four new size/Settings/diagnostic checks passed
while implementing the corrections. The final build repeats every suite, the
unintended-startup-window check, native helper checks, frozen Excel checks and
actual-payload installation/uninstallation in owned temporary directories.
Completed packaged results/checksums are recorded in the final release-folder copy.
Embedded app documentation is a preflight snapshot, not certification of a later EXE.

## Authorized cleanup

Retain source and required assets/native resources, scripts/tests, documentation,
one latest dist/FINDEX executable folder, one canonical local installer release
and the prepared github/FINDEX source tree. After verification remove obsolete
releases/builds/caches/EXEs/backups and generated Metadata_Reports, as requested.
Remove local virtual environments last. Running source again requires recreating
an environment from requirements; the standalone app and installer bundle their
runtime and helper tools. Original .py source and Git metadata are preserved.
No source evidence or files outside this project are eligible for cleanup.

## Distribution limits

The prior native review is a historical record dated 2026-11-14. Native dependency
versions are unchanged; unresolved native findings/corresponding-source conditions
are not claimed cleared. The prepared GitHub tree excludes native binaries. Keep
this installer/application local until its separate native distribution gates are
completed. The installer remains unsigned. Functional tests and a local antivirus
scan do not certify forensic reliability, absence of vulnerabilities or every PC's
compatibility. Clean-VM/native compliance work remains separate.

## Completed isolated-build checks — 2026-11-22

- All 60 regression tests passed in the stable temporary C: workspace, including
  the real Windows junction-pruning case; no tests were skipped in that run.
- No unexpected top-level window appeared before the main window in the startup check.
- Actual frozen EXE: seven real Excel/HTML/hash reports generated and reopened
  successfully, both normally and with a simulated incomplete legacy NumPy module.
- Actual installed payload: integrity/extraction and both frozen Excel cases passed
  in a Unicode installation path. ExifTool 13.59 and ffprobe 9.0.2 processed the
  synthetic Unicode evidence without changing its content or modification time.
- Uninstall retained the user report and evidence fixture; measured local duration 5.42s.
- Setup EXE identity, icon and asInvoker checks passed. Setup startup passed offscreen;
  Install was not pressed and no shortcuts/Installed Apps entries were created by testing.
- The release verifier checked every source-archive inventory/hash and all 875
  application payload files before handoff. This is not forensic certification.
- The 17 locked Python packages passed a refreshed pip-audit check with no known
  package advisories reported. That check does not clear the historical native findings.
- Initial E: runs were interrupted by device-not-present/write errors. The source
  snapshot/build/tests were completed in an isolated C: folder before final delivery.
  No obsolete release was deleted while a verified replacement was unavailable.

Final endpoint scan, delivery hash verification and cleanup are recorded in the
release-folder copy after those steps. Local installer/application use remains
separate from the source-only GitHub publication and its native distribution limits.

## Final handoff — 2026-11-22

The permanent corrected project is C:\Users\Arshad\Downloads\FINDEX.
The final retry of E: failed with Windows device-not-present error 433. The old
E: project, releases and generated reports were left untouched; cleanup there is
deferred until that drive is reliable. No verified replacement depends on E:.

The clean C: delivery retains source/resources, one complete dist/FINDEX app,
one release installer and github/FINDEX source. It excludes temporary builds,
virtual environments, generated reports and redundant release ZIP copies.
Original source copies on E: remain preserved. Recreate a Python environment from
requirements to run source; the standalone application bundles its runtime,
openpyxl, ExifTool and ffprobe and needs no separate installation of those tools.

The executable at the permanent C: destination passed its seven-report Excel/HTML
check with external Python/virtual-environment paths removed. See EXCEL_CHECK.json
in the release folder. The final local app/installer Defender scan reported zero
matching detections (signature 1.463.403.0); this is a scan result, not certification.
Source and application inventories and release checksums are verified at handoff.

The GitHub source repository is initialized on main. No commit, remote upload,
GitHub release or LinkedIn post was performed. Upload the prepared source tree;
keep the bundled application and installer local pending the native distribution
conditions described above and in PUBLIC_RELEASE.md.
