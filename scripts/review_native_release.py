"""Record the reviewed local binary's native inventory and open distribution issues."""
import ast
import hashlib
import json
import subprocess
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def digest(stream):
    result = hashlib.sha256()
    for block in iter(lambda: stream.read(1024 * 1024), b''):
        result.update(block)
    return result.hexdigest()


def main():
    release = ROOT / 'release/FINDEX-v3.0.1'
    portable = release / 'FINDEX_v3.0.1_Portable.zip'
    manifest = json.loads((release / 'RELEASE_MANIFEST.json').read_text(encoding='utf-8'))
    expected = next(r['sha256'] for r in manifest['files'] if r['path'] == portable.name)
    with portable.open('rb') as stream:
        assert digest(stream) == expected, 'Portable archive no longer matches reviewed manifest'
    identifiers = set()
    source_hashes = {}
    for name in ('FINDEX.py', 'installer_wizard.py', 'dialog_theme.py', 'report_branding.py'):
        data = (ROOT / name).read_bytes()
        source_hashes[name] = hashlib.sha256(data).hexdigest()
        for node in ast.walk(ast.parse(data)):
            if isinstance(node, ast.Name):
                identifiers.add(node.id)
            elif isinstance(node, ast.Attribute):
                identifiers.add(node.attr)
            elif isinstance(node, ast.ImportFrom) and node.module and node.module.startswith('PySide6'):
                identifiers.update(alias.name for alias in node.names)
    rows = [
        ('CVE-2026-116337', 'Qt Network HTTP responses', ('Qt6Network.dll',), ('QNetworkAccessManager', 'QNetworkReply')),
        ('CVE-2026-110944', 'Qt HTTP Server', ('Qt6HttpServer.dll',), ('QHttpServer',)),
        ('CVE-2026-114366', 'Qt WebSockets', ('Qt6WebSockets.dll',), ('QWebSocket',)),
        ('CVE-2026-111487', 'Qt3D assimp importer', ('Qt63DCore.dll', 'Qt63DRender.dll'), ('QSceneLoader',)),
        ('CVE-2026-100140', 'Qt5Compat XML', ('Qt6Core5Compat.dll',), ('QXmlSimpleReader',)),
        ('CVE-2026-105924', 'Qt Core CBOR conversion', ('Qt6Core.dll',), ('QCborValue', 'QCborStreamReader')),
    ]
    inventory = []
    with zipfile.ZipFile(portable) as archive:
        names = archive.namelist()
        dll_names = {Path(name).name for name in names if '/PySide6/' in name and name.endswith('.dll')}
        for name in names:
            if '/PySide6/' in name and (Path(name).name.startswith('Qt6') and name.endswith('.dll') or name.endswith('.pyd')):
                with archive.open(name) as stream:
                    inventory.append(dict(path=name, size=archive.getinfo(name).file_size, sha256=digest(stream)))
    findings = []
    for cve, component, modules, apis in rows:
        present = sorted(set(modules) & dll_names)
        direct = sorted(set(apis) & identifiers)
        assert not direct, 'Application API use changed; reassess ' + cve
        findings.append(dict(cve=cve, component=component, checked_modules=modules,
            present_modules=present, direct_api_references=direct,
            disposition='present_no_direct_trigger_found_not_cleared' if present else 'triggering_module_not_shipped'))
    metadata = json.loads((ROOT / 'scratch/native_release_metadata.json').read_text(encoding='utf-8'))
    probe = ROOT / 'external/ffmpeg/ffprobe.exe'
    with probe.open('rb') as stream:
        probe_hash = digest(stream)
    version = subprocess.run([str(probe), '-version'], capture_output=True, text=True,
                             check=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW).stdout
    readme = (ROOT / 'external/ffmpeg/README.txt').read_text(encoding='utf-8')
    versions = readme.split("release-essentials external libraries' versions:", 1)[1].strip().splitlines()
    report = dict(reviewed_on='2026-11-14', scope='Existing local 3.0.1 portable; source-only GitHub preparation',
        portable_sha256=expected, application_sources=source_hashes,
        qt_version='6.12.0', qt_advisory_source='https://wiki.qt.io/List_of_known_vulnerabilities_in_Qt_products',
        qt_inventory=inventory, qt_findings=findings, public_package_metadata=metadata,
        ffmpeg=dict(binary_sha256=probe_hash, version_and_configuration=version,
                    provenance=json.loads((ROOT / 'external/ffmpeg/SOURCE.json').read_text()),
                    recorded_external_library_versions=versions,
                    complete_corresponding_source_verified=False,
                    missing_materials=['Exact external dependency sources and local patches',
                        'Build recipes/configuration inputs and corresponding source completeness verification',
                        'Source coverage for dependencies absent from the recorded library-version list']),
        source_publication='No native binaries or runtime libraries included; original FINDEX source remains MIT',
        binary_publication='ON HOLD: native redistribution materials and present Qt findings remain unresolved',
        limits='AST references and artifact inventory are not a full transitive native reachability or exploitability audit.')
    (ROOT / 'NATIVE_DISTRIBUTION_REVIEW.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    table = '\n'.join('| ' + f['cve'] + ' | ' + f['component'] + ' | ' +
        ('Present; no direct trigger found; remains open' if f['present_modules'] else 'Triggering module absent from artifact') + ' |' for f in findings)
    text = f'''# Native distribution review — FINDEX 3.0.1

Reviewed 2026-11-14. Application logic, UI, branding and native dependencies were not changed.

## Decision

Prepare the original FINDEX source for GitHub, without native binaries. Keep the
existing Setup EXE and portable ZIP out of the public commit/downloads. Personal
or noncommercial intent does not remove binary redistribution conditions.
This review does not certify legal compliance or absence of vulnerabilities.

## Qt advisory review

The exact portable archive has SHA-256 `{expected}`. Its native Qt inventory and
application-source hashes are in NATIVE_DISTRIBUTION_REVIEW.json. The reviewed
advisory findings relevant to Qt 6.12.0 are:

| Advisory | Triggering component | Local artifact disposition |
|---|---|---|
{table}

Source analysis found no direct application references to the checked triggering
APIs. That does not prove all native/transitive paths unreachable. Core and
Network findings remain open; their presence is not cleared by a pip-audit result.
The four absent-module dispositions apply to this exact artifact, not every Qt installation.

PyPI metadata currently lists PySide6 {metadata['pyside6_latest']} as latest.
6.12.1 is unavailable. 6.11.3 exists, but the cited HTTP-response advisory does
not list that branch as fixed; a downgrade does not establish complete resolution.
Keep the tested pin, disclose the findings, and re-evaluate compatible patches.
Reference: https://wiki.qt.io/List_of_known_vulnerabilities_in_Qt_products

## FFmpeg distribution review

FINDEX uses ffprobe as a separate executable for media duration. The current
Gyan 9.0.2 build is static and enables GPL/version3 components. The binary hash,
configuration, upstream revision, archive provenance and {len(versions)} recorded
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
'''
    (ROOT / 'NATIVE_DISTRIBUTION_REVIEW.md').write_text(text, encoding='utf-8')
    print('Reviewed', len(inventory), 'Qt binaries/bindings; findings:', len(findings))
    print('FFmpeg source completeness is unresolved; source-only publication excludes native binaries.')


if __name__ == '__main__':
    main()
