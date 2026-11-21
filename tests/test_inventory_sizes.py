"""Report size units, Settings persistence and actual Excel workbooks."""
import hashlib
import json
import subprocess
import sys
import unittest
from unittest.mock import patch
import test_review_fixes as fixtures
from test_review_fixes import m, wait_for


class InventorySizeTests(unittest.TestCase):
    def setUp(self):
        self.case = fixtures.ReviewFixTests('test_all_category_extensions_remain_reachable')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)
        self.addCleanup(self.case.tearDown)
        self.root, self.source = self.case.root, self.case.source

    def test_worker_reports_all_units_and_preserves_source(self):
        large = self.source / 'large.bin'
        large.write_bytes(b'x' * 1572864)
        before = (hashlib.sha256(large.read_bytes()).hexdigest(), large.stat().st_mtime_ns)
        cases = {'auto': ('3 B', '1.50 MB', 'File Size'),
                 'kb': ('0.00 KB', '1536.00 KB', 'File Size (KB)'),
                 'mb': ('0.00 MB', '1.50 MB', 'File Size (MB)')}
        for unit, (small_value, large_value, header) in cases.items():
            with self.subTest(unit=unit):
                excel = self.root / f'{unit}.xlsx'
                html = self.root / f'{unit}.html'
                done, errors, cancelled, _ = self.case.inventory(
                    size_unit=unit, enabled_columns=['si', 'name', 'size'],
                    output_excel=str(excel), output_html=str(html))
                self.assertFalse(errors)
                self.assertFalse(cancelled)
                self.assertEqual(done[0][1]['total_size'], 1572867)
                workbook = m.openpyxl.load_workbook(excel)
                try:
                    rows = [row for row in workbook.active if isinstance(row[0].value, int)]
                    values = {row[1].value: row[2].value for row in rows}
                    self.assertEqual(values, {'sample.txt': small_value, 'large.bin': large_value})
                    self.assertIn(header, [cell.value for row in workbook.active for cell in row])
                finally:
                    workbook.close()
                rendered = html.read_text(encoding='utf-8')
                for expected in (small_value, large_value, header):
                    self.assertIn(expected, rendered)
        self.assertEqual(m.COLUMN_MAP['size']['name'], 'File Size')
        self.assertEqual((hashlib.sha256(large.read_bytes()).hexdigest(), large.stat().st_mtime_ns), before)

    def test_settings_persist_feed_worker_and_reset(self):
        window = self.case.make_window()
        try:
            window.settings_inventory_size_combo.setCurrentIndex(1)
            window._save_workstation_settings()
            self.assertEqual(window.settings_store.value('inventory_size_unit'), 'kb')
        finally:
            window.close()
        window = self.case.make_window()
        try:
            self.assertEqual(window.settings_inventory_size_combo.currentData(), 'kb')
            window.inv_src_edit.setText(str(self.source))
            window.inv_out_edit.setText(str(self.root / 'settings-size'))
            window.rb_both.setChecked(True)
            with patch.object(m.CompletionDialog, 'exec', return_value=0):
                window._start_inventory_scan()
                self.assertEqual(window.inventory_worker.config['size_unit'], 'kb')
                wait_for(window.inventory_worker)
            window._reset_workstation_settings()
            self.assertEqual(window.settings_inventory_size_combo.currentData(), 'auto')
            self.assertEqual(window.settings_store.value('inventory_size_unit', 'auto'), 'auto')
        finally:
            window.close()

    def test_invalid_columns_have_consistent_defaults(self):
        folder = [dict(folder='MAIN', files=[dict(si=1, name='synthetic.txt')], total_size=0)]
        output = self.root / 'empty-columns.xlsx'
        m.ExcelReportGenerator.generate(output, folder, [], 'Synthetic', 1, 0)
        workbook = m.openpyxl.load_workbook(output)
        try:
            rows = list(workbook.active.values)
            self.assertIn(('SI No.', 'File Name'), rows)
            self.assertIn((1, 'synthetic.txt'), rows)
        finally:
            workbook.close()

    def test_source_diagnostic_exports_real_workbooks(self):
        output = self.root / 'excel-check'
        checked = subprocess.run([sys.executable, str(fixtures.ROOT / 'File Inventory Pro.py'),
            '--check-excel', str(output)], cwd=fixtures.ROOT, capture_output=True, timeout=45)
        result = json.loads((output / 'result.json').read_text(encoding='utf-8'))
        self.assertEqual(checked.returncode, 0, result)
        self.assertEqual(result['status'], 'passed')
        self.assertEqual(len(result['reports']), 7)


if __name__ == '__main__':
    unittest.main(verbosity=2)
