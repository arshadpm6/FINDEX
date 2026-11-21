"""Overwrite consent, metadata report formatting and verification UI regressions."""
import copy
import json
import queue
import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch
import test_review_fixes as fixtures
from test_review_fixes import m, app, wait_for
from security_io import (new_report, write_report, report_write_scope, report_identity,
                         report_key, ReportSaveCancelled)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QMessageBox


class GuiCorrectionsTests(unittest.TestCase):
    def setUp(self):
        self.case = fixtures.ReviewFixTests('test_all_category_extensions_remain_reachable')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)
        self.addCleanup(self.case.tearDown)
        self.root, self.source = self.case.root, self.case.source

    def window(self):
        window = self.case.make_window()
        self.addCleanup(window.close)
        return window

    def test_replace_requires_consent_and_preserves_original_on_failure(self):
        target = self.root / 'report.txt'
        target.write_text('old report', encoding='utf-8')
        with report_write_scope(authorize=lambda _: False):
            with self.assertRaises(ReportSaveCancelled):
                write_report(target, 'new')
        self.assertEqual(target.read_text(), 'old report')
        with report_write_scope(authorize=lambda _: True):
            with self.assertRaises(RuntimeError):
                with new_report(target) as output:
                    output.write('partial new')
                    raise RuntimeError('simulated disk/writer error')
        self.assertEqual(target.read_text(), 'old report')
        self.assertFalse(list(self.root.glob('.findex-report-*')))
        with report_write_scope(authorize=lambda _: True):
            write_report(target, 'replacement complete')
        self.assertEqual(target.read_text(), 'replacement complete')

    def test_approved_target_cannot_change_before_replacement(self):
        target = self.root / 'report.txt'
        target.write_text('old')
        approved = {report_key(target): report_identity(target)}
        target.write_text('changed outside FINDEX')
        with report_write_scope(approved=approved):
            with self.assertRaises(OSError):
                write_report(target, 'new')
        self.assertEqual(target.read_text(), 'changed outside FINDEX')

    def test_evidence_never_replaced_even_with_consent(self):
        sample = self.case.sample
        for options in (dict(protected_paths=[sample]), dict(protected_roots=[self.source])):
            with report_write_scope(authorize=lambda _: True, **options):
                with self.assertRaises(ValueError):
                    write_report(sample, 'report')
        self.assertEqual(sample.read_bytes(), b'abc')

    def test_hash_save_dialog_yes_no_and_no_duplicate_prompt(self):
        window = self.window()
        window._add_file_to_hash_items(str(self.case.sample))
        target = self.root / 'hashes.csv'
        target.write_text('existing report')
        with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')) as dialog:
            with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.No) as question:
                window._export_hash_csv()
                self.assertEqual(target.read_text(), 'existing report')
                question.assert_called_once()
            with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.Yes) as question:
                window._export_hash_csv()
                self.assertIn('FileName', target.read_text())
                question.assert_called_once()
            self.assertEqual(dialog.call_args.kwargs['options'], QFileDialog.Option.DontConfirmOverwrite)

    def test_inventory_replacement_approved_on_gui_before_worker(self):
        window = self.window()
        window.inv_src_edit.setText(str(self.source))
        window.inv_out_edit.setText(str(self.root / 'inventory'))
        window.rb_both.setChecked(True)
        window.cb_docs.setChecked(True)
        for extension in ('xlsx', 'html'):
            (self.root / ('inventory.' + extension)).write_text('old')
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.No):
            window._start_inventory_scan()
        self.assertEqual((self.root / 'inventory.html').read_text(), 'old')
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.Yes), patch.object(m.CompletionDialog, 'exec', return_value=0):
            window._start_inventory_scan()
            wait_for(window.inventory_worker)
        self.assertTrue((self.root / 'inventory.xlsx').read_bytes().startswith(b'PK'))
        self.assertIn('<html', (self.root / 'inventory.html').read_text(encoding='utf-8'))

    def test_metadata_all_formats_existing_folder_and_raw_sidecar(self):
        window = self.window()
        record = self.case.record()
        window.active_meta_record = record
        window.all_meta_records = [record]
        window.meta_cb_report_html.setChecked(True)
        window.meta_cb_report_txt.setChecked(True)
        window.meta_cb_report_json.setChecked(True)
        window.meta_cb_preserve_json.setChecked(True)
        directory = self.root / 'exports'
        directory.mkdir()
        unrelated = directory / 'personal.txt'
        unrelated.write_text('retain')
        with patch.object(QFileDialog, 'getExistingDirectory', return_value=str(directory)), patch.object(QMessageBox, 'question') as confirmation:
            window._export_metadata_by_selected_format()
            confirmation.assert_not_called()
        with patch.object(QFileDialog, 'getExistingDirectory', return_value=str(directory)), patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.Yes) as confirmation:
            window._export_metadata_by_selected_format()
            self.assertEqual(confirmation.call_count, 4)
        self.assertEqual(len(list(directory.iterdir())), 5)
        self.assertTrue(list(directory.glob('*_Metadata.txt')))
        self.assertEqual(unrelated.read_text(), 'retain')
        raw = next(directory.glob('*_exiftool_raw.json'))
        self.assertEqual(raw.read_text(), record['raw_json'])

    def test_print_style_and_long_structured_content(self):
        record = self.case.record()
        record['categories'] = {'Other Metadata': [dict(tag='Long structure', raw_key='Other:Long', value={'values': list(range(300))})]}
        target = self.root / 'metadata.html'
        m.MetadataReportExporter.export_html(target, record)
        text = target.read_text(encoding='utf-8')
        for rule in ('size: A4 portrait', 'margin: 14mm 12mm 15mm', 'table-layout: fixed',
                     'display: table-header-group', 'max-height: none', 'overflow: visible'):
            self.assertIn(rule, text)
        self.assertIn('299', text)
        self.assertIn('class="structured-row"', text)
        self.assertIn("class='struct-line'", text)

    def test_metadata_txt_save_as_replacement(self):
        window = self.window()
        record = self.case.record()
        window.active_meta_record = record
        window.all_meta_records = [record]
        window.meta_cb_report_html.setChecked(False)
        window.meta_cb_report_txt.setChecked(True)
        window.meta_cb_report_json.setChecked(False)
        window.meta_cb_preserve_json.setChecked(False)
        target = self.root / 'metadata.txt'
        target.write_text('old TXT report')
        with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')):
            with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.No):
                window._export_metadata_by_selected_format()
            self.assertEqual(target.read_text(), 'old TXT report')
            with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.Yes) as confirmation:
                window._export_metadata_by_selected_format()
                confirmation.assert_called_once()
        text = target.read_text(encoding='utf-8')
        self.assertIn('ExifTool', text)
        self.assertIn(record['file_name'], text)
        self.assertEqual(window._last_meta_txt, str(target))

    def test_animation_elapsed_time_no_resets_with_progress(self):
        loader = m.AnimatedVerificationLoader()
        self.addCleanup(loader.close)
        loader.timer.stop()
        with patch.object(m.time, 'monotonic', return_value=100.0):
            loader.set_state('verifying')
        with patch.object(m.time, 'monotonic', return_value=100.125):
            loader._on_tick()
        self.assertAlmostEqual(loader.angle, 11.25)
        with patch.object(m.time, 'monotonic', return_value=101.0):
            for count in range(1000):
                loader.set_state('verifying', f'File {count}')
            loader._on_tick()
        self.assertEqual(loader._verification_started_at, 100.0)
        self.assertAlmostEqual(loader.angle, 90.0)
        self.assertEqual(loader.timer.timerType(), Qt.TimerType.PreciseTimer)

    def test_large_result_queue_does_not_rebuild_table_per_file(self):
        window = self.window()
        count = 1500
        window.hash_items = [dict(id=str(i), name=f'{i}.txt', path=str(self.case.sample), size=3,
            algorithm='SHA-256', expected_hash='', result='PENDING', status_text='Pending') for i in range(count)]
        window._refresh_hash_table()
        window._hash_models_by_id = {item['id']: item for item in window.hash_items}
        results = queue.SimpleQueue()
        for item in window.hash_items:
            update = dict(item, result='NO TARGET', status='computed', status_text='Calculated', calculated_hash='a' * 64)
            results.put((item['id'], update))
        window.hash_worker = SimpleNamespace(ui_results=results, ui_progress=None, isRunning=lambda: False)
        with patch.object(window, '_refresh_hash_table') as full_rebuild:
            window._on_hash_completed(dict(total=count, matched=0, mismatched=0, computed=count, errors=0))
            self.assertGreater(results.qsize(), 0)
            while not results.empty():
                window._drain_hash_updates()
            full_rebuild.assert_not_called()
        self.assertTrue(all(item['result'] == 'NO TARGET' for item in window.hash_items))
        self.assertEqual(window.hash_table.item(count - 1, 6).text(), 'a' * 64)
        self.assertEqual(window.kpi_unchecked.value_lbl.text(), f'{count:,}')


if __name__ == '__main__':
    unittest.main(verbosity=2)
