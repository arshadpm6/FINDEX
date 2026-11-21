"""Security regression tests with disposable, harmless fixtures only."""
import csv
import hashlib
import io
import json
import os
import subprocess
import sys
import unittest
import zipfile
from unittest.mock import patch
from test_review_fixes import ReviewFixTests, m, installer, ROOT, QMessageBox
from security_io import new_report, csv_text, script_string, evidence_walk
from release_integrity import payload_manifest
from PIL import Image


class SecurityTests(ReviewFixTests):
    def test_incompatible_numpy_does_not_block_application_startup(self):
        script = """
import importlib.util
import sys
import types
from pathlib import Path

root = Path.cwd()
sys.path.insert(0, str(root))
sys.modules["numpy"] = types.ModuleType("numpy")
spec = importlib.util.spec_from_file_location("findex_numpy_test", root / "File Inventory Pro.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
assert module.OPENPYXL_AVAILABLE
assert module.OPENPYXL_IMPORT_ERROR is None
assert "numpy" in module.OPENPYXL_NUMPY_FALLBACK
assert "short" in module.OPENPYXL_NUMPY_FALLBACK
assert sys.modules["numpy"].__name__ == "numpy"
import tempfile
from excel_diagnostics import check_excel_exports
with tempfile.TemporaryDirectory(dir=root / "scratch") as temporary:
    assert check_excel_exports(module, Path(temporary) / "excel-check") == 0
"""
        environment = os.environ.copy()
        environment["QT_QPA_PLATFORM"] = "offscreen"
        result = subprocess.run(
            [sys.executable, "-c", script],
            cwd=ROOT,
            env=environment,
            capture_output=True,
            text=True,
            timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_report_does_not_substitute_another_algorithm(self):
        item = dict(name='sample.txt', path=str(self.sample), size=3, algorithm='MD5',
                    all_hashes={'SHA-256': 'cached-other-algorithm'}, calculated_hash='', result='PENDING')
        output = self.root / 'hashes.csv'
        m.HashReportExporter.export_csv(output, [item], default_algorithm='SHA-256')
        row = list(csv.reader(io.StringIO(output.read_text(encoding='utf-8'))))[1]
        self.assertEqual(row[4], 'MD5')
        self.assertEqual(row[6], '')
        self.assertNotIn('cached-other-algorithm', output.read_text())
    def test_branding_and_themes_preserved(self):
        window = self.make_window()
        try:
            self.assertEqual(window.windowTitle(), 'FINDEX \u2013 File Intelligence & Examination v3.0.1')
            self.assertFalse(window.windowIcon().isNull())
            self.assertEqual(window.sidebar_brand_title.text(), 'FINDEX')
            self.assertEqual(window.sidebar_brand_subtitle.text(), 'File Intelligence & Examination')
            for theme in ('dark', 'light'):
                window._apply_global_theme(theme)
                about = m.AboutFindexDialog(window)
                self.assertFalse(about.windowIcon().isNull())
                about.close()
                for switch, expected in ((window._switch_to_hash_mode, 'hash'), (window._switch_to_inventory_mode, 'inventory'), (window._switch_to_metadata_mode, 'metadata')):
                    switch()
                    self.assertEqual(window.active_mode, expected)
        finally:
            window.close()

    def test_metadata_hash_failure_is_visible(self):
        records, summaries = [], []
        worker = m.MetadataExtractionWorker([str(self.sample)], {'include_hashes': True})
        worker.file_completed.connect(lambda _, record, *args: records.append(record))
        worker.extraction_completed.connect(summaries.append)
        with patch.object(m.MetadataHashAdapter, 'calculate_integrity', side_effect=OSError('fixture locked')):
            worker.run()
        self.assertEqual(records[0]['integrity_info']['status'], 'HASH_FAILED')
        self.assertTrue(any('fixture locked' in message for message in records[0]['warnings']))
        self.assertGreater(summaries[0]['warnings'], 0)

    def test_report_cannot_overwrite_evidence_and_partial_cleanup(self):
        original = self.sample.read_bytes()
        for exporter in (m.HashReportExporter.export_csv, m.HashReportExporter.export_html,
                         m.HashReportExporter.export_excel, m.HashReportExporter.export_checksum_file):
            with self.assertRaises(FileExistsError):
                exporter(self.sample, [])
            self.assertEqual(self.sample.read_bytes(), original)
        partial = self.root / 'partial.txt'
        with self.assertRaises(RuntimeError):
            with new_report(partial) as output:
                output.write('partial')
                raise RuntimeError('fixture failure')
        self.assertFalse(partial.exists())
        with self.assertRaises(ValueError):
            with new_report(str(self.sample) + ':hidden'):
                self.fail('Alternate data stream must not be created')

    def test_link_pruning_callback(self):
        with patch('security_io.os.walk', return_value=[(str(self.source), ['linked-dir'], ['linked-file', 'sample.txt'])]), patch('security_io.linked', side_effect=lambda p: p.name.startswith('linked')):
            warnings = []
            entries = list(evidence_walk(self.source, onerror=warnings.append))
            self.assertEqual(entries[0][1:], ([], ['sample.txt']))
            self.assertEqual(len(warnings), 2)

    def test_exiftool_timeout_and_malformed_output(self):
        with patch.object(m.subprocess, 'run', side_effect=subprocess.TimeoutExpired('fixture', 1)):
            record = m.ExifToolEngine.extract_metadata(self.sample, timeout=1)
            self.assertFalse(record['success'])
            self.assertIn('timed out', record['error'])
        with patch.object(m.subprocess, 'run', return_value=subprocess.CompletedProcess([], 0, '{bad JSON', '')):
            self.assertFalse(m.ExifToolEngine.extract_metadata(self.sample)['success'])

    def test_streaming_large_file_and_source_preservation(self):
        target = self.source / 'large file é नमूना.bin'
        expected = hashlib.sha256()
        block = bytes(range(256)) * 4096
        with target.open('wb') as output:
            for _ in range(32):
                output.write(block)
                expected.update(block)
        before = target.stat()
        progress = []
        actual = m.compute_all_file_hashes(target, algorithms={'SHA-256'}, chunk_size=65536,
                                          chunk_callback=progress.append)
        self.assertEqual(actual['SHA-256'], expected.hexdigest())
        self.assertEqual(len(progress), 512)
        self.assertEqual(target.stat().st_mtime_ns, before.st_mtime_ns)
        self.assertEqual(target.stat().st_size, before.st_size)

    def test_change_during_hash_is_rejected(self):
        def change(count):
            if count == 1:
                self.sample.write_bytes(b'changed fixture')
        with self.assertRaisesRegex(OSError, 'changed while hashing'):
            m.compute_all_file_hashes(self.sample, chunk_size=1, chunk_callback=change)

    def test_expected_hash_validation(self):
        for invalid in ('z' * 64, 'a' * 32, '=malicious'):
            results = []
            item = dict(id='test', path=str(self.sample), name='sample.txt', algorithm='SHA-256', expected_hash=invalid)
            worker = m.HashVerificationWorker([item], 'SHA-256')
            worker.item_completed.connect(lambda _, value: results.append(value))
            worker.run()
            self.assertEqual(results[0]['result'], 'ERROR')
            self.assertFalse(results[0]['calculated_hash'])

    def test_csv_and_script_encoding(self):
        rows = list(csv.reader(io.StringIO(csv_text([['=1+1', 'a,"b\nc', '@SUM(A1)', 'é']]))))
        self.assertEqual(rows[0], ["'=1+1", 'a,"b\nc', "'@SUM(A1)", 'é'])
        hostile = "</script><script>alert('x')</script>\u2028"
        encoded = script_string(hostile)
        self.assertNotIn('<', encoded)
        self.assertEqual(json.loads(encoded), hostile)

    def test_metadata_html_escape_and_evidence_unchanged(self):
        target = self.source / 'é नमूना image.png'
        Image.new('RGB', (2, 2)).save(target)
        before = (target.read_bytes(), target.stat().st_mtime_ns)
        record = m.ExifToolEngine.extract_metadata(target)
        self.assertTrue(record['success'], record)
        hostile = '<script>alert("x")</script>'
        record['file_name'] = hostile
        record['categories'] = {'Other Metadata': [{'tag': hostile, 'value': hostile, 'key': hostile}]}
        output = self.root / 'hostile.html'
        m.MetadataReportExporter.export_html(output, record)
        result = output.read_text(encoding='utf-8')
        self.assertNotIn(hostile, result)
        self.assertIn('&lt;script&gt;', result)
        self.assertEqual((target.read_bytes(), target.stat().st_mtime_ns), before)

    def test_exiftool_does_not_load_local_config(self):
        config = self.root / '.ExifTool_config'
        config.write_text('die "UNSAFE CONFIG LOADED";', encoding='utf-8')
        previous = os.getcwd()
        try:
            os.chdir(self.root)
            version, success, _ = m.ExifToolEngine.get_version()
            self.assertTrue(success)
            self.assertEqual(version, '13.59')
        finally:
            os.chdir(previous)

    def test_junction_is_pruned(self):
        junction = self.source / 'loop'
        self.assertTrue(junction.absolute().is_relative_to(ROOT / 'scratch'))
        command = "New-Item -ItemType Junction -Path '" + str(junction).replace("'", "''") + "' -Target '" + str(self.source).replace("'", "''") + "' | Out-Null"
        result = subprocess.run(['powershell', '-NoProfile', '-Command', command], capture_output=True)
        if result.returncode != 0:
            self.skipTest('Workspace filesystem does not support creating Windows junctions')
        try:
            warnings = []
            walked = list(evidence_walk(self.source, onerror=warnings.append))
            self.assertEqual(len(walked), 1)
            self.assertEqual(len(warnings), 1)
        finally:
            junction.rmdir()  # Removes the fixture junction itself, never its target.

    def test_executable_open_requires_confirmation(self):
        target = self.root / 'harmless.exe'
        target.write_bytes(b'fixture, not executable')
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.No), patch.object(m.QDesktopServices, 'openUrl') as opened:
            self.assertFalse(m.open_evidence_file(None, target))
            opened.assert_not_called()

    def test_installer_integrity_traversal_and_collision(self):
        payload = self.root / 'payload.zip'
        def create(name, data=b'fixture'):
            with zipfile.ZipFile(payload, 'w') as archive:
                archive.writestr('FINDEX.exe', b'app')
                archive.writestr(name, data)
            with zipfile.ZipFile(payload) as archive:
                manifest = payload_manifest(archive)
            payload.with_suffix('.manifest.json').write_text(json.dumps(manifest), encoding='utf-8')
        dest = self.root / 'FINDEX'
        create('personal.txt')
        dest.mkdir()
        (dest / 'FINDEX.exe').write_bytes(b'old')
        (dest / 'personal.txt').write_bytes(b'preserve')
        for tamper in (False, True):
            if tamper:
                with zipfile.ZipFile(payload, 'a') as archive:
                    archive.writestr('unexpected.txt', b'tampered')
            results = []
            worker = installer.ExtractWorker(payload, dest, False, False, False)
            worker.finished.connect(lambda *args: results.append(args))
            worker.run()
            self.assertFalse(results[0][0], results)
            self.assertEqual((dest / 'FINDEX.exe').read_bytes(), b'old')
            self.assertEqual((dest / 'personal.txt').read_bytes(), b'preserve')
        for unsafe in ('../outside.txt', '/absolute.txt', 'x:stream', 'CON', 'folder/../x', 'x. '):
            with self.assertRaises(ValueError):
                create(unsafe)


del ReviewFixTests  # The inherited regression tests run once, in SecurityTests.

if __name__ == '__main__':
    unittest.main(verbosity=2)
