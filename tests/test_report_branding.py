"""FINDEX identity, self-contained icons and stable logo backing in reports/UI."""
import base64
import hashlib
import re
import unittest
from html.parser import HTMLParser
from pathlib import Path
import test_review_fixes as fixtures
from PySide6.QtGui import QPalette, QFontDatabase, QFont

m, app = fixtures.m, fixtures.app


class HtmlElements(HTMLParser):
    def __init__(self, text):
        super().__init__(convert_charrefs=True)
        self.elements = []
        self.feed(text)

    def handle_starttag(self, tag, attributes):
        self.elements.append((tag, dict(attributes)))


class ReportBrandingTests(unittest.TestCase):
    def setUp(self):
        self.case = fixtures.ReviewFixTests('test_all_category_extensions_remain_reachable')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)
        self.addCleanup(self.case.tearDown)
        self.root = self.case.root
        for name in ('segoeui.ttf', 'segoeuib.ttf', 'seguisb.ttf'):
            QFontDatabase.addApplicationFont(str(Path('C:/Windows/Fonts') / name))
        app.setFont(QFont('Segoe UI', 10))

    def reports(self):
        inventory = self.root / 'inventory.html'
        data = [dict(folder='MAIN FOLDER', total_size=3, files=[dict(si=1, name='sample.txt',
            size='3 B', format='TXT', path=str(self.case.sample), sha256=hashlib.sha256(b'abc').hexdigest())])]
        m.HtmlReportGenerator.generate(inventory, data, ['si', 'name', 'size', 'format', 'sha256'],
                                       str(self.case.source), 1, 3)
        hashes = self.root / 'hashes.html'
        m.HashReportExporter.export_html(hashes, [dict(name='sample.txt', path=str(self.case.sample),
            size=3, algorithm='SHA-256', result='NO TARGET', status='computed', status_text='Calculated',
            calculated_hash=hashlib.sha256(b'abc').hexdigest())])
        metadata = self.root / 'metadata.html'
        m.MetadataReportExporter.export_html(metadata, self.case.record())
        return [inventory, hashes, metadata]

    def test_all_report_types_use_findex_name_version_and_new_logo(self):
        expected_logo = (fixtures.ROOT / 'resources/branding/report_logo.png').read_bytes()
        for path in self.reports():
            with self.subTest(report=path.name):
                text = path.read_text(encoding='utf-8')
                parsed = HtmlElements(text)
                self.assertIn('<title>FINDEX — ', text)
                self.assertIn('File Intelligence &amp; Examination', text)
                self.assertIn('v3.0.1', text)
                self.assertIn('Know Every File.', text)
                self.assertNotIn('File Inventory Pro', text)
                self.assertNotIn('PRO REPORT', text)
                logos = [attributes for tag, attributes in parsed.elements if tag == 'img']
                self.assertEqual(len(logos), 1)
                self.assertEqual(logos[0]['alt'], 'FINDEX logo')
                self.assertEqual(base64.b64decode(logos[0]['src'].split(',', 1)[1]), expected_logo)
                favicons = [attributes for tag, attributes in parsed.elements if tag == 'link' and attributes.get('rel') == 'icon']
                self.assertEqual(favicons[0]['href'], logos[0]['src'])
                self.assertNotRegex(text, r'[\U0001f300-\U0001faff\u2600-\u27bf]')

    def test_legacy_icon_cache_and_desktop_logo_are_untouched(self):
        desktop = fixtures.ROOT / 'resources/branding/favicon.png'
        legacy = fixtures.ROOT / 'assets/favicon.png'
        before = {path: hashlib.sha256(path.read_bytes()).digest() for path in (desktop, legacy) if path.is_file()}
        self.reports()
        for path, digest in before.items():
            self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), digest)
        self.assertLess((fixtures.ROOT / 'resources/branding/report_logo.png').stat().st_size, 50000)

    def test_inventory_actions_have_vector_icons_and_theme_toggle_keeps_them(self):
        text = self.reports()[0].read_text(encoding='utf-8')
        parsed = HtmlElements(text)
        vectors = [attributes for tag, attributes in parsed.elements if tag == 'svg']
        self.assertGreaterEqual(len(vectors), 7)
        self.assertTrue(all(attributes.get('aria-hidden') == 'true' for attributes in vectors))
        self.assertIn('themeToggleLabel', text)
        self.assertIn("getElementById('themeToggleLabel').textContent", text)
        self.assertNotIn("getElementById('themeToggleBtn').innerText", text)
        self.assertIn('onclick="exportToCsv()"', text)
        self.assertIn('onclick="window.print()"', text)
        self.assertIn('oninput="handleSearch()"', text)
        self.assertIn('<h1 class="report-brand-name">FINDEX', text)

    def test_sidebar_logo_has_same_dark_square_and_original_dimensions(self):
        window = self.case.make_window()
        try:
            window.resize(1366, 900)
            window.show()
            original = None
            output = fixtures.ROOT / 'scratch/report_branding_preview'
            output.mkdir(exist_ok=True)
            for theme in ('dark', 'light'):
                window._apply_global_theme(theme)
                app.processEvents()
                logo = window.sidebar_brand_logo
                self.assertEqual((logo.width(), logo.height()), (30, 30))
                self.assertEqual(window.sidebar_frame.width(), 230)
                self.assertEqual(logo.palette().color(QPalette.Window).name(), '#111625')
                image = logo.pixmap().toImage()
                if original is not None:
                    self.assertEqual(image, original)
                original = image
                label = window.sidebar_brand_subtitle
                self.assertLessEqual(label.fontMetrics().horizontalAdvance(label.text()), label.width())
                window.sidebar_frame.grab().save(str(output / ('sidebar_' + theme + '.png')))
        finally:
            window.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
