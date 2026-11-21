"""Prompt only for actual file replacement; export only selected metadata formats."""
import copy
import json
import unittest
from unittest.mock import patch
import test_review_fixes as fixtures
from PySide6.QtWidgets import QFileDialog, QMessageBox

m, app = fixtures.m, fixtures.app


class ExportSelectionTests(unittest.TestCase):
    def setUp(self):
        self.case = fixtures.ReviewFixTests('test_all_category_extensions_remain_reachable')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)
        self.addCleanup(self.case.tearDown)
        self.root = self.case.root
        self.directory = self.root / 'existing exports'
        self.directory.mkdir()

    def window(self, html=False, txt=False, json_format=False, raw=False):
        window = self.case.make_window()
        self.addCleanup(window.close)
        record = self.case.record()
        window.active_meta_record = record
        window.all_meta_records = [record]
        window.meta_cb_report_html.setChecked(html)
        window.meta_cb_report_txt.setChecked(txt)
        window.meta_cb_report_json.setChecked(json_format)
        window.meta_cb_preserve_json.setChecked(raw)
        return window, record

    def test_json_and_raw_are_opt_in(self):
        window = self.case.make_window()
        self.addCleanup(window.close)
        self.assertFalse(window.meta_cb_report_json.isChecked())
        self.assertFalse(window.meta_cb_preserve_json.isChecked())
        self.assertFalse(window.meta_cb_preserve_json.isEnabled())
        window.meta_cb_report_json.setChecked(True)
        self.assertTrue(window.meta_cb_preserve_json.isEnabled())
        window.meta_cb_preserve_json.setChecked(True)
        self.assertTrue(window._metadata_raw_export_requested())
        window.meta_cb_report_json.setChecked(False)
        self.assertFalse(window.meta_cb_preserve_json.isEnabled())
        self.assertFalse(window._metadata_raw_export_requested())

    def test_html_and_txt_only_do_not_create_json_or_prompt_for_existing_folder(self):
        for format_name in ('html', 'txt'):
            with self.subTest(format=format_name):
                # Simulate a previously checked raw option after JSON was deselected.
                window, _ = self.window(html=format_name == 'html', txt=format_name == 'txt', raw=True)
                target = self.directory / ('single.' + format_name)
                with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')):
                    with patch.object(QMessageBox, 'question') as question:
                        window._export_metadata_by_selected_format()
                        question.assert_not_called()
                self.assertTrue(target.is_file())
                self.assertFalse(list(self.directory.glob('*.json')))

    def test_multiple_formats_reuse_folder_and_preserve_unselected_old_json(self):
        window, record = self.window(html=True, txt=True, raw=True)
        basename = m.metadata_report_basename(record)
        old_raw = self.directory / (basename + '_exiftool_raw.json')
        old_raw.write_text('{"previous_export": true}', encoding='utf-8')
        unrelated = self.directory / 'personal.txt'
        unrelated.write_text('keep this')
        before = old_raw.read_bytes(), old_raw.stat().st_mtime_ns
        with patch.object(QFileDialog, 'getExistingDirectory', return_value=str(self.directory)):
            with patch.object(QMessageBox, 'question') as question:
                window._export_metadata_by_selected_format()
                question.assert_not_called()
        self.assertEqual({p.name for p in self.directory.iterdir()}, {
            old_raw.name, 'personal.txt', basename + '_Metadata.html', basename + '_Metadata.txt'})
        self.assertEqual((old_raw.read_bytes(), old_raw.stat().st_mtime_ns), before)
        self.assertEqual(unrelated.read_text(), 'keep this')
        html = (self.directory / (basename + '_Metadata.html')).read_text(encoding='utf-8')
        self.assertNotIn('fetch(', html)
        self.assertIn('FINDEX', html)

    def test_actual_metadata_filename_collision_decline_then_replace(self):
        window, _ = self.window(html=True)
        target = self.directory / 'report.html'
        target.write_text('original report')
        with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')):
            with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.No) as question:
                window._export_metadata_by_selected_format()
                question.assert_called_once()
                self.assertEqual(question.call_args.args[1], 'Replace Existing File?')
                self.assertIn(str(target), question.call_args.args[2])
            self.assertEqual(target.read_text(), 'original report')
            with patch.object(QMessageBox, 'question', return_value=QMessageBox.StandardButton.Yes) as question:
                window._export_metadata_by_selected_format()
                question.assert_called_once()
        self.assertIn('FINDEX', target.read_text(encoding='utf-8'))
        self.assertFalse(list(self.directory.glob('*.json')))

    def test_same_filename_elsewhere_is_not_a_collision(self):
        window, _ = self.window()
        window._add_file_to_hash_items(str(self.case.sample))
        other = self.root / 'another folder'
        other.mkdir()
        (other / 'hashes.csv').write_text('a different file')
        target = self.directory / 'hashes.csv'
        with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')):
            with patch.object(QMessageBox, 'question') as question:
                window._export_hash_csv()
                question.assert_not_called()
        self.assertTrue(target.is_file())
        self.assertEqual((other / 'hashes.csv').read_text(), 'a different file')

    def test_json_export_saves_raw_only_when_explicitly_selected(self):
        for raw in (False, True):
            with self.subTest(raw=raw):
                window, record = self.window(json_format=True, raw=raw)
                output = self.directory / ('with raw' if raw else 'record only')
                output.mkdir()
                target = output / 'selected_metadata.json'
                with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')):
                    with patch.object(QMessageBox, 'question') as question:
                        window._export_metadata_by_selected_format()
                        question.assert_not_called()
                files = list(output.iterdir())
                self.assertEqual(len(files), 2 if raw else 1)
                payload = json.loads(target.read_text(encoding='utf-8'))
                self.assertEqual(payload['raw_json'], record['raw_json'] if raw else None)
                if raw:
                    self.assertEqual((output / 'selected_exiftool_raw.json').read_text(), record['raw_json'])

    def test_inventory_new_file_in_existing_folder_needs_no_replacement_prompt(self):
        window, _ = self.window()
        window.inv_src_edit.setText(str(self.case.source))
        window.inv_out_edit.setText(str(self.directory / 'inventory'))
        window.rb_both.setChecked(True)
        window.cb_docs.setChecked(True)
        with patch.object(QMessageBox, 'question') as question, patch.object(m.CompletionDialog, 'exec', return_value=0):
            window._start_inventory_scan()
            fixtures.wait_for(window.inventory_worker)
            question.assert_not_called()
        self.assertTrue((self.directory / 'inventory.html').is_file())
        self.assertTrue((self.directory / 'inventory.xlsx').is_file())

    def test_inventory_prefix_does_not_prompt_for_unselected_extension(self):
        window, _ = self.window()
        window.inv_src_edit.setText(str(self.case.source))
        window.rb_excel.setChecked(True)
        window.cb_docs.setChecked(True)
        existing_html = self.directory / 'prefix.html'
        existing_html.write_text('Unselected HTML must be retained')
        with patch.object(QFileDialog, 'getSaveFileName', return_value=(str(existing_html), '')) as dialog:
            window._browse_output_file()
            self.assertEqual(dialog.call_args.kwargs['options'], QFileDialog.Option.DontConfirmOverwrite)
        with patch.object(QMessageBox, 'question') as question, patch.object(m.CompletionDialog, 'exec', return_value=0):
            window._start_inventory_scan()
            fixtures.wait_for(window.inventory_worker)
            question.assert_not_called()
        self.assertTrue((self.directory / 'prefix.xlsx').is_file())
        self.assertEqual(existing_html.read_text(), 'Unselected HTML must be retained')

    def test_worker_auto_export_respects_json_selection(self):
        for json_format, raw in ((False, True), (True, False), (True, True)):
            with self.subTest(json=json_format, raw=raw):
                directory = self.directory / f'worker_{json_format}_{raw}'
                record = self.case.record()
                options = dict(auto_export=True, output_dir=str(directory), include_hashes=False,
                    gen_html=not json_format, gen_txt=False, gen_json=json_format, preserve_raw_json=raw)
                worker = m.MetadataExtractionWorker([str(self.case.sample)], options)
                summaries = []
                worker.extraction_completed.connect(summaries.append)
                with patch.object(m.ExifToolEngine, 'extract_metadata', return_value=copy.deepcopy(record)):
                    worker.run()
                self.assertEqual(summaries[0]['successful'], 1)
                self.assertEqual(len(list(directory.glob('*.json'))), (2 if raw else 1) if json_format else 0)
                self.assertEqual(len(list(directory.glob('*.html'))), 0 if json_format else 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
