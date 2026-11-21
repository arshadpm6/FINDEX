"""Theme standard FINDEX message boxes independently of Windows' dialog palette."""
from PySide6.QtCore import QObject, QEvent
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QMessageBox


class MessageBoxThemeFilter(QObject):
    def eventFilter(self, watched, event):
        if not isinstance(watched, QMessageBox) or QApplication.closingDown():
            return False
        if event.type() in (
                QEvent.Type.Polish, QEvent.Type.Show,
                QEvent.Type.StyleChange, QEvent.Type.PaletteChange):
            parent = watched.parentWidget()
            theme = None
            while parent is not None:
                if getattr(parent, 'current_theme', None) in ('light', 'dark'):
                    theme = parent.current_theme
                    break
                parent = parent.parentWidget()
            if theme is None:
                theme = QApplication.instance().property('findexTheme') or 'dark'
            if watched.property('findexMessageTheme') != theme:
                # Palette/style updates generate events; guard against recursive styling.
                watched.setProperty('findexMessageTheme', theme)
                self.apply_theme(watched, theme)
            if event.type() == QEvent.Type.Show:
                for button in watched.buttons():
                    button.ensurePolished()
                    text_width = button.fontMetrics().horizontalAdvance(button.text().replace('&', ''))
                    button.setMinimumWidth(max(96, text_width + 36))
                watched.layout().activate()
                watched.adjustSize()
        # QObject's default filter does nothing; do not delegate through a wrapper
        # that may already be tearing down during Qt/Python shutdown.
        return False

    @staticmethod
    def apply_theme(dialog, theme):
        light = theme == 'light'
        background = '#ffffff' if light else '#161b2b'
        foreground = '#0f172a' if light else '#f1f5f9'
        secondary = '#475569' if light else '#aab6c8'
        button = '#f1f5f9' if light else '#20273d'
        hover = '#e2e8f0' if light else '#2d3852'
        pressed = '#cbd5e1' if light else '#334155'
        border = '#94a3b8' if light else '#64748b'
        accent, accent_hover = '#2563eb', '#1d4ed8'
        # Apply to message boxes only; native file dialogs keep their existing behavior.
        dialog.setOption(QMessageBox.Option.DontUseNativeDialog, True)
        palette = dialog.palette()
        for role, color in (
                (QPalette.Window, background), (QPalette.WindowText, foreground),
                (QPalette.Base, background), (QPalette.AlternateBase, button),
                (QPalette.Text, foreground), (QPalette.Button, button),
                (QPalette.ButtonText, foreground), (QPalette.ToolTipBase, background),
                (QPalette.ToolTipText, foreground), (QPalette.Highlight, accent),
                (QPalette.HighlightedText, '#ffffff'),
                (QPalette.Link, '#1d4ed8' if light else '#93c5fd')):
            palette.setColor(role, QColor(color))
        for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
            palette.setColor(QPalette.Disabled, role, QColor(secondary))
        dialog.setPalette(palette)
        dialog.setStyleSheet(f"""
            QMessageBox {{ background-color: {background}; color: {foreground}; }}
            QMessageBox QLabel {{
                color: {foreground}; background: transparent; border: none;
                font-family: 'Segoe UI'; font-size: 13px;
            }}
            QMessageBox QLabel#qt_msgbox_informativelabel {{ color: {secondary}; }}
            QMessageBox QDialogButtonBox {{ background: transparent; border: none; }}
            QMessageBox QPushButton {{
                background-color: {button}; color: {foreground};
                border: 1px solid {border}; border-radius: 6px;
                padding: 7px 16px; min-width: 64px;
                font-family: 'Segoe UI'; font-size: 13px;
            }}
            QMessageBox QPushButton:hover {{ background-color: {hover}; }}
            QMessageBox QPushButton:pressed {{ background-color: {pressed}; }}
            QMessageBox QPushButton:focus {{ border: 2px solid {accent}; padding: 6px 15px; }}
            QMessageBox QPushButton:default {{
                background-color: {accent}; color: #ffffff; border-color: {accent};
            }}
            QMessageBox QPushButton:default:hover {{ background-color: {accent_hover}; }}
            QMessageBox QPushButton:disabled {{
                background-color: {button}; color: {secondary}; border-color: {border};
            }}
            QMessageBox QTextEdit, QMessageBox QPlainTextEdit {{
                background-color: {background}; color: {foreground};
                border: 1px solid {border}; selection-background-color: {accent};
                selection-color: #ffffff;
            }}
            QMessageBox QCheckBox {{ color: {foreground}; background: transparent; }}
        """)


def install_message_box_theme_filter():
    app = QApplication.instance()
    if app is not None and not hasattr(app, '_findex_message_box_filter'):
        app._findex_message_box_filter = MessageBoxThemeFilter(app)
        app.installEventFilter(app._findex_message_box_filter)
