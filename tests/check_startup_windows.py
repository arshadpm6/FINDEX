"""Detect unintended top-level widgets becoming visible during construction."""
import os
import runpy
from PySide6.QtCore import QObject, QEvent
from PySide6.QtWidgets import QWidget

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
ns = runpy.run_path('tests/test_review_fixes.py', run_name='startup_check')
app = ns['app']

class VisibilityMonitor(QObject):
    def __init__(self):
        super().__init__()
        self.windows = []

    def eventFilter(self, target, event):
        if event.type() == QEvent.Type.Show and isinstance(target, QWidget) and target.isWindow():
            self.windows.append(type(target).__name__)
        return False

case = ns['ReviewFixTests']('test_actual_window_inventory_worker_and_ui_completion')
case.setUp()
monitor = VisibilityMonitor()
app.installEventFilter(monitor)
try:
    window = case.make_window()
    print('Top-level windows shown before main window:', monitor.windows)
    assert not monitor.windows, monitor.windows
    window.show()
    app.processEvents()
    print('Top-level windows after main show:', monitor.windows)
    assert monitor.windows == ['FileInventoryProApp'], monitor.windows
    window.close()
finally:
    app.removeEventFilter(monitor)
    case.tearDown()
    case.doCleanups()
