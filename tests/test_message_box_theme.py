"""Contrast checks for real warning/confirmation dialogs in both FINDEX themes."""
import unittest
from pathlib import Path
import test_review_fixes as fixtures
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QColor, QPalette, QFont, QFontDatabase
from PySide6.QtWidgets import QMessageBox, QLabel, QDialog, QTextEdit, QApplication

m, app = fixtures.m, fixtures.app


def luminance(color):
    channels = [component / 255 for component in (color.red(), color.green(), color.blue())]
    linear = [c / 12.92 if c <= .04045 else ((c + .055) / 1.055) ** 2.4 for c in channels]
    return sum(c * weight for c, weight in zip(linear, (.2126, .7152, .0722)))


def contrast(first, second):
    values = sorted((luminance(first), luminance(second)))
    return (values[1] + .05) / (values[0] + .05)


class MessageBoxThemeTests(unittest.TestCase):
    def setUp(self):
        self.case = fixtures.ReviewFixTests('test_all_category_extensions_remain_reachable')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)
        self.addCleanup(self.case.tearDown)
        self.old_palette = app.palette()
        self.addCleanup(app.setPalette, self.old_palette)
        for filename in ('segoeui.ttf', 'segoeuib.ttf'):
            QFontDatabase.addApplicationFont(str(Path('C:/Windows/Fonts') / filename))
        app.setFont(QFont('Segoe UI', 10))
        self.window = self.case.make_window()
        self.addCleanup(self.window.close)
        self.output = fixtures.ROOT / 'scratch/message_box_previews'
        self.output.mkdir(parents=True, exist_ok=True)

    def opposite_windows_palette(self, theme):
        palette = app.palette()
        palette.setColor(QPalette.Window, QColor('#202020' if theme == 'light' else '#ffffff'))
        palette.setColor(QPalette.WindowText, QColor('#ffffff' if theme == 'light' else '#000000'))
        palette.setColor(QPalette.Text, QColor('#ffffff' if theme == 'light' else '#000000'))
        app.setPalette(palette)

    def check_contrast(self, box, theme):
        app.processEvents()
        self.assertEqual(box.property('findexMessageTheme'), theme)
        self.assertTrue(box.testOption(QMessageBox.Option.DontUseNativeDialog))
        background = box.palette().color(QPalette.Window)
        self.assertEqual(background.name(), '#ffffff' if theme == 'light' else '#161b2b')
        for label in box.findChildren(QLabel):
            if label.text():
                foreground = label.palette().color(label.foregroundRole())
                self.assertGreaterEqual(contrast(foreground, background), 4.5, label.text())
        for button in box.buttons():
            palette = button.palette()
            self.assertGreaterEqual(contrast(palette.color(QPalette.ButtonText), palette.color(QPalette.Button)), 4.5)
            self.assertGreaterEqual(button.width(), button.fontMetrics().horizontalAdvance(button.text().replace('&', '')) + 20)
        for editor in box.findChildren(QTextEdit):
            palette = editor.palette()
            self.assertGreaterEqual(contrast(palette.color(QPalette.Text), palette.color(QPalette.Base)), 4.5)

    def check_all_icons(self, theme):
        self.opposite_windows_palette(theme)
        self.window._apply_global_theme(theme)
        for name, icon in (('warning', QMessageBox.Icon.Warning), ('question', QMessageBox.Icon.Question),
                           ('information', QMessageBox.Icon.Information), ('critical', QMessageBox.Icon.Critical)):
            with self.subTest(theme=theme, type=name):
                box = QMessageBox(icon, 'FINDEX - ' + name.title(),
                    'The selected report already exists.\nReplace it with the new report?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, self.window)
                box.setInformativeText('Source evidence is protected. Choosing No preserves the existing file.')
                box.setDetailedText('Diagnostic details remain readable in the selected application theme.')
                box.setDefaultButton(QMessageBox.StandardButton.No)
                box.show()
                self.check_contrast(box, theme)
                self.assertTrue(box.button(QMessageBox.StandardButton.No).isDefault())
                box.grab().save(str(self.output / (theme + '_' + name + '.png')))
                for button in box.buttons():
                    if box.buttonRole(button) == QMessageBox.ButtonRole.ActionRole:
                        button.click()
                        break
                self.check_contrast(box, theme)
                box.close()

    def test_light_dialogs_when_windows_is_dark(self):
        self.check_all_icons('light')

    def test_dark_dialogs_when_windows_is_light(self):
        self.check_all_icons('dark')

    def test_file_confirmation_actual_modal_default_and_result(self):
        self.opposite_windows_palette('light')
        self.window._apply_global_theme('light')
        captured, errors = [], []
        def respond():
            box = app.activeModalWidget()
            if not isinstance(box, QMessageBox):
                errors.append('No message box was shown')
                return
            try:
                self.check_contrast(box, 'light')
                self.assertTrue(box.button(QMessageBox.StandardButton.No).isDefault())
                box.grab().save(str(self.output / 'light_file_confirmation.png'))
                captured.append(box.windowTitle())
            except Exception as error:
                errors.append(str(error))
            finally:
                box.button(QMessageBox.StandardButton.No).click()
        QTimer.singleShot(50, respond)
        result = self.window._confirm_report_replace('E:/test reports/report.html')
        self.assertFalse(result)
        self.assertFalse(errors, errors)
        self.assertEqual(captured, ['Replace Existing File?'])

    def test_nested_dialog_parent_does_not_override_application_theme(self):
        self.window._apply_global_theme('light')
        parent = QDialog(self.window)
        parent.setStyleSheet('QDialog {background:#202020;} QLabel {color:#000000;}')
        box = QMessageBox(QMessageBox.Icon.Warning, 'FINDEX', 'Check the selected path.',
                          QMessageBox.StandardButton.Ok, parent)
        box.show()
        self.check_contrast(box, 'light')
        box.close()
        parent.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
