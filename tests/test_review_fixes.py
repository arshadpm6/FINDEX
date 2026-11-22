"""Regression checks for the review fixes, using the actual Qt application."""
import hashlib
import gc
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
import wave
import zipfile
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


m = load('inventory_app_tests', 'FINDEX.py')
installer = load('inventory_installer_tests', 'installer_wizard.py')
from PySide6.QtCore import QSettings, QTimer
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

app = QApplication.instance() or QApplication([])


def wait_for(worker, timeout=20):
    deadline = time.monotonic() + timeout
    while worker.isRunning() and time.monotonic() < deadline:
        app.processEvents()
        time.sleep(.01)
    if worker.isRunning():
        worker.cancel()
        worker.wait(5000)
        raise AssertionError('Worker did not finish in time')
    app.processEvents()


class ReviewFixTests(unittest.TestCase):
    def setUp(self):
        (ROOT / 'scratch').mkdir(exist_ok=True)
        for name in ('information', 'warning', 'critical'):
            dialog_patch = patch.object(QMessageBox, name, return_value=QMessageBox.StandardButton.Ok)
            dialog_patch.start()
            self.addCleanup(dialog_patch.stop)
        self.temp = tempfile.TemporaryDirectory(prefix='fixed_test_', dir=ROOT / 'scratch')
        self.root = Path(self.temp.name)
        self.source = self.root / 'evidence'
        self.source.mkdir()
        self.sample = self.source / 'sample.txt'
        self.sample.write_bytes(b'abc')
        m.HASH_CHUNK_SIZE = 65536
        m.ENABLED_HASH_ALGORITHMS = set(m.HASH_ALGORITHMS)

    def tearDown(self):
        app.processEvents()
        gc.collect()
        if not Path(self.temp.name).resolve().is_relative_to((ROOT / 'scratch').resolve()):
            raise AssertionError('Refusing cleanup outside the owned test workspace')
        # Windows may briefly retain delete-pending files/directories after the
        # fixture's uninstaller exits or a scanner releases a freshly created ZIP.
        # Bound retries to the owned fixture and transient Windows errors only.
        for attempt in range(20):
            try:
                self.temp.cleanup()
                break
            except OSError as error:
                if getattr(error, 'winerror', None) not in (5, 32, 145) or attempt == 19:
                    raise
                time.sleep(.25)

    def inventory(self, **overrides):
        config = dict(source_folder=str(self.source), output_excel=str(self.root / 'report.xlsx'),
                      output_html=str(self.root / 'report.html'), format_type='both',
                      enabled_columns=['si', 'name', 'size', 'format', 'created', 'modified', 'path', 'md5', 'sha1', 'sha256'],
                      filter_categories=dict(audio=True, video=True, images=True, docs=True, archives=True, other=True),
                      si_mode='continuous', skip_empty_folders=False)
        config.update(overrides)
        worker = m.InventoryWorker(config)
        done, errors, cancelled, no_files = [], [], [], []
        worker.inventory_finished.connect(lambda *args: done.append(args))
        worker.error_occurred.connect(lambda *args: errors.append(args))
        worker.inventory_cancelled.connect(lambda *args: cancelled.append(args))
        worker.no_files_found.connect(lambda *args: no_files.append(args))
        worker.start()
        wait_for(worker)
        return done, errors, cancelled, no_files

    def test_inventory_end_to_end_excel_html_hashes_and_numbering(self):
        sub = self.source / 'nested'
        sub.mkdir()
        (sub / 'video.mpeg').write_bytes(b'video')
        (sub / 'voice.opus').write_bytes(b'audio')
        (sub / '=1+1.txt').write_bytes(b'formula filename')
        (self.source / 'empty').mkdir()
        done, errors, cancelled, _ = self.inventory()
        self.assertFalse(errors)
        self.assertFalse(cancelled)
        self.assertEqual(len(done), 1)
        paths, stats, warnings = done[0]
        self.assertEqual(stats['total_files'], 4)
        self.assertEqual(stats['total_folders'], 3)
        self.assertEqual(stats['total_size'], sum(p.stat().st_size for p in self.source.rglob('*') if p.is_file()))
        self.assertFalse(warnings)
        workbook = m.openpyxl.load_workbook(paths[0])
        rows = [row for row in workbook.active.iter_rows() if isinstance(row[0].value, int)]
        self.assertEqual([row[0].value for row in rows], [1, 2, 3, 4])
        by_name = {row[1].value: row for row in rows}
        self.assertEqual(by_name['=1+1.txt'][1].data_type, 's')
        self.assertEqual(by_name['sample.txt'][-1].value, hashlib.sha256(b'abc').hexdigest())
        self.assertTrue(by_name['sample.txt'][4].value)
        workbook.close()
        html = Path(paths[1]).read_text(encoding='utf-8')
        for name in ('sample.txt', 'video.mpeg', 'voice.opus', '=1+1.txt'):
            self.assertIn(name, html)
        self.assertIn(hashlib.sha256(b'abc').hexdigest(), html)

    def test_inventory_folder_serials_skip_empty_and_filter(self):
        sub = self.source / 'nested'
        sub.mkdir()
        (sub / 'voice.opus').write_bytes(b'voice')
        (self.source / 'empty').mkdir()
        done, errors, _, _ = self.inventory(si_mode='folder', skip_empty_folders=True)
        self.assertFalse(errors)
        self.assertEqual(done[0][1]['total_folders'], 2)
        wb = m.openpyxl.load_workbook(self.root / 'report.xlsx')
        self.assertEqual([row[0].value for row in wb.active.iter_rows() if isinstance(row[0].value, int)], [1, 1])
        wb.close()
        done, errors, _, _ = self.inventory(filter_categories={'audio': True},
            output_excel=str(self.root / 'audio.xlsx'), output_html=str(self.root / 'audio.html'))
        self.assertFalse(errors)
        self.assertEqual(done[0][1]['total_files'], 1)

    def test_inventory_real_media_duration(self):
        with wave.open(str(self.source / 'one_second.wav'), 'wb') as audio:
            audio.setnchannels(1)
            audio.setsampwidth(2)
            audio.setframerate(8000)
            audio.writeframes(b'\x00\x00' * 8000)
        done, errors, _, _ = self.inventory(filter_categories={'audio': True}, enabled_columns=['si', 'name', 'duration'])
        self.assertFalse(errors)
        self.assertFalse(done[0][2])
        wb = m.openpyxl.load_workbook(self.root / 'report.xlsx')
        row = next(row for row in wb.active.iter_rows() if row[1].value == 'one_second.wav')
        self.assertEqual(row[2].value, '00:00:01')
        wb.close()

    def test_inventory_cancel_no_files_and_unexpected_failure(self):
        cancelled = []
        cancelled_worker = m.InventoryWorker(dict(source_folder=str(self.source), format_type='html',
            filter_categories={'docs': True}, enabled_columns=['si', 'name'], si_mode='folder'))
        cancelled_worker.inventory_cancelled.connect(cancelled.append)
        cancelled_worker.cancel()
        cancelled_worker.run()
        self.assertEqual(cancelled, [0])
        worker = m.InventoryWorker({})
        errors = []
        worker.error_occurred.connect(lambda *args: errors.append(args))
        worker.run()
        self.assertEqual(len(errors), 1)
        done, errors, _, no_files = self.inventory(filter_categories={'images': True})
        self.assertFalse(done)
        self.assertFalse(errors)
        self.assertEqual(len(no_files), 1)
        with patch.object(m.InventoryWorker, '_run_inventory', side_effect=RuntimeError('injected failure')):
            done, errors, _, _ = self.inventory()
        self.assertFalse(done)
        self.assertIn('injected failure', errors[0][1])

    def test_all_category_extensions_remain_reachable(self):
        allowed = m.AUDIO_EXTENSIONS | m.VIDEO_EXTENSIONS | m.IMAGE_EXTENSIONS | m.DOCUMENT_EXTENSIONS | m.ARCHIVE_EXTENSIONS
        self.assertEqual(allowed, m.ALL_NAMED_EXTENSIONS)

    def record(self, path=None):
        return dict(success=True, file_name='sample.txt', file_path=str(path or self.sample), file_size=3,
                    parsed={'File:FileName': 'sample.txt'}, categories={}, raw_json='[ { "raw": "preserve me" } ]')

    def test_exiftool_real_error_and_success(self):
        bad = self.source / 'unsupported.bin'
        bad.write_bytes(b'unknown\x00metadata\x01file')
        failed = m.ExifToolEngine.extract_metadata(bad)
        self.assertFalse(failed['success'])
        self.assertIn('Unknown file type', failed['error'])
        from PIL import Image
        image = self.source / 'sample_photo_01.jpg'
        Image.new('RGB', (3, 3), 'blue').save(image)
        good = m.ExifToolEngine.extract_metadata(image)
        self.assertTrue(good['success'], good.get('error'))
        self.assertTrue(good['categories'])

    def test_metadata_collisions_and_pre_hash_only_status(self):
        records = [self.record(), self.record(self.root / 'other' / 'sample.txt')]
        paths = [m.MetadataReportExporter.export_json(self.root / 'reports', rec) for rec in records]
        self.assertNotEqual(paths[0][0], paths[1][0])
        self.assertEqual(len(list((self.root / 'reports').iterdir())), 4)
        info = dict(pre=m.compute_all_file_hashes(self.sample), post={}, verified=None)
        m.MetadataReportExporter.export_html(self.root / 'pre.html', records[0], info)
        m.MetadataReportExporter.export_txt(self.root / 'pre.txt', records[0], info)
        self.assertNotIn('Mismatch Detected', (self.root / 'pre.html').read_text(encoding='utf-8'))
        self.assertIn('NOT COMPARED', (self.root / 'pre.txt').read_text(encoding='utf-8'))

    def test_hash_failure_invalidates_old_result_and_manifest_roundtrip(self):
        item = dict(id='one', path=str(self.sample), name='sample.txt', expected_hash=hashlib.sha256(b'abc').hexdigest(), algorithm='SHA-256')
        results = []
        worker = m.HashVerificationWorker([item], 'SHA-256')
        worker.item_completed.connect(lambda _, value: results.append(value))
        worker.run()
        good = results[-1]
        self.assertEqual(good['result'], 'MATCH')
        manifest = self.root / 'elsewhere' / 'checksums.sha256'
        m.HashReportExporter.export_checksum_file(manifest, [good])
        entry = manifest.read_text().split(None, 1)[1].strip()
        self.assertEqual((manifest.parent / entry).resolve(), self.sample.resolve())
        good['path'] = str(self.root / 'missing.txt')
        worker = m.HashVerificationWorker([good], 'SHA-256')
        worker.item_completed.connect(lambda _, value: results.append(value))
        worker.run()
        failed = results[-1]
        self.assertEqual(failed['result'], 'ERROR')
        self.assertFalse(failed['calculated_hash'])
        self.assertFalse(failed['all_hashes'])
        failed_manifest = manifest.with_name('failed.sha256')
        m.HashReportExporter.export_checksum_file(failed_manifest, [failed])
        self.assertEqual(failed_manifest.read_text(), '')

    def test_worker_snapshots_and_metadata_cancellation(self):
        items = [dict(id=str(i), path=str(self.sample), name='sample.txt', algorithm='SHA-256') for i in range(2)]
        worker = m.HashVerificationWorker(items, 'SHA-256')
        items.clear()
        summaries = []
        worker.verification_completed.connect(summaries.append)
        worker.run()
        self.assertEqual(summaries[0]['computed'], 2)
        metadata = m.MetadataExtractionWorker([str(self.sample)], dict(include_hashes=True, verify_integrity=True))
        cancelled, done = [], []
        metadata.extraction_cancelled.connect(cancelled.append)
        metadata.extraction_completed.connect(done.append)
        def cancel_hash(*args, **kwargs):
            metadata.cancel()
            return None
        with patch.object(m.MetadataHashAdapter, 'calculate_integrity', side_effect=cancel_hash), \
             patch.object(m.ExifToolEngine, 'extract_metadata') as extract:
            metadata.run()
        self.assertEqual(cancelled, [0])
        self.assertFalse(done)
        extract.assert_not_called()

    def make_window(self):
        settings = QSettings(str(self.root / 'settings.ini'), QSettings.Format.IniFormat)
        settings.setValue('auto_open_folder', False)
        with patch.object(m, 'QSettings', return_value=settings):
            window = m.FileInventoryProApp()
        return window

    def test_actual_window_inventory_worker_and_ui_completion(self):
        window = self.make_window()
        try:
            window.inv_src_edit.setText(str(self.source))
            window.inv_out_edit.setText(str(self.root / 'gui_report'))
            window.rb_both.setChecked(True)
            window.cb_docs.setChecked(True)
            window.col_checkboxes['duration'].setChecked(False)
            with patch.object(m.CompletionDialog, 'exec', return_value=0), patch.object(QMessageBox, 'warning') as warning:
                window._start_inventory_scan()
                wait_for(window.inventory_worker)
            self.assertTrue(window.btn_run_inventory.isEnabled())
            self.assertIn('Complete', window.inv_status_lbl.text())
            self.assertEqual(len(window.last_generated_files), 2)
            self.assertTrue((self.root / 'gui_report.xlsx').is_file())
            self.assertTrue((self.root / 'gui_report.html').is_file())
            warning.assert_not_called()
        finally:
            window.close()

    def test_actual_window_shortcuts_settings_summary_and_raw_export(self):
        window = self.make_window()
        try:
            with patch.object(QFileDialog, 'getOpenFileName', return_value=(str(self.sample), '')):
                window._quick_action_metadata_file()
            self.assertEqual(window.meta_src_edit.text(), str(self.sample))
            self.assertEqual(window.active_mode, 'metadata')
            with patch.object(m.QDesktopServices, 'openUrl') as opened:
                window._show_preview_tab()
            opened.assert_called_once()
            with patch.object(window.verification_loader, 'set_state') as state:
                window._on_hash_completed(dict(total=2, matched=1, mismatched=0, computed=0, errors=1))
            self.assertNotEqual(state.call_args.args[0], 'match')
            record = self.record()
            window.meta_cb_report_json.setChecked(True)
            window.meta_cb_preserve_json.setChecked(True)
            target = self.root / 'manual.json'
            window._write_meta_json_record(target, record)
            self.assertEqual(json.loads(target.read_text())['raw_json'], record['raw_json'])
            self.assertEqual((self.root / 'manual_exiftool_raw.json').read_text(), record['raw_json'])
            window.settings_chunk_combo.setCurrentIndex(3)
            window.settings_cb_md5.setChecked(False)
            window.settings_output_edit.setText(str(self.root / 'chosen_output'))
            window._save_workstation_settings()
            self.assertEqual(m.HASH_CHUNK_SIZE, 1048576)
            self.assertNotIn('MD5', m.compute_all_file_hashes(self.sample))
            self.assertEqual(Path(window.inv_out_edit.text()).parent, self.root / 'chosen_output')
        finally:
            window.close()

    def test_comparison_uses_worker(self):
        dialog = m.CompareTwoFilesDialog(str(self.sample), str(self.sample))
        try:
            dialog._do_compare()
            self.assertFalse(dialog.compare_btn.isEnabled())
            wait_for(dialog.comparison_worker)
            self.assertIn('MATCH CONFIRMED', dialog.result_banner.text())
            self.assertTrue(dialog.compare_btn.isEnabled())
        finally:
            dialog.close()

    def test_installer_rejects_shared_destination_and_records_files(self):
        with self.assertRaises(ValueError):
            installer.validate_destination(self.root)
        dest = self.root / 'File Inventory Pro'
        dest.mkdir()
        unrelated = dest / 'personal.txt'
        unrelated.write_text('retain')
        with self.assertRaises(ValueError):
            installer.validate_destination(dest)
        # A recognized application folder can be upgraded; preserve unrelated contents.
        (dest / 'File Inventory Pro.exe').write_bytes(b'old')
        payload = self.root / 'payload.zip'
        with zipfile.ZipFile(payload, 'w') as archive:
            archive.writestr('FINDEX.exe', b'new')
            archive.writestr('_internal/data.txt', b'installed')
        from release_integrity import payload_manifest
        with zipfile.ZipFile(payload) as archive:
            payload.with_suffix('.manifest.json').write_text(json.dumps(payload_manifest(archive)), encoding='utf-8')
        worker = installer.ExtractWorker(payload, dest, False, False, False)
        outcomes = []
        worker.finished.connect(lambda *args: outcomes.append(args))
        worker.run()
        self.assertTrue(outcomes[0][0], outcomes)
        manifest = json.loads((dest / 'install_manifest.json').read_text())
        self.assertNotIn('personal.txt', manifest['files'])
        self.assertEqual(unrelated.read_text(), 'retain')
        script = (dest / 'uninstall.ps1').read_text(encoding='utf-8-sig')
        self.assertNotIn('rd /s', script)
        self.assertNotIn('-Recurse', script)
        self.assertIn('install_manifest.json', script)
        # This fixture has no shortcut/registry options and only workspace-local files.
        self.assertTrue(dest.resolve().is_relative_to(ROOT / 'scratch'))
        completed = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                                    '-File', str(dest / 'uninstall.ps1')], capture_output=True, text=True)
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(unrelated.read_text(), 'retain')
        self.assertFalse((dest / 'FINDEX.exe').exists())
        self.assertFalse((dest / '_internal' / 'data.txt').exists())


if __name__ == '__main__':
    import faulthandler
    faulthandler.dump_traceback_later(30, repeat=True)
    unittest.main(verbosity=2)
