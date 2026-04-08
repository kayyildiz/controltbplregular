
import sys
import types


class _Dummy:
    def __init__(self, *args, **kwargs):
        pass

    def __call__(self, *args, **kwargs):
        return _Dummy()

    def __getattr__(self, name):
        return _Dummy()

    def __iter__(self):
        return iter(())

    def __bool__(self):
        return False

    def __lt__(self, other):
        return False


class _Qt:
    ControlModifier = 0
    Key_V = 0
    WaitCursor = 0


def install():
    if 'PySide6' in sys.modules:
        return

    pyside6 = types.ModuleType('PySide6')
    qtcore = types.ModuleType('PySide6.QtCore')
    qtgui = types.ModuleType('PySide6.QtGui')
    qtwidgets = types.ModuleType('PySide6.QtWidgets')

    class QTimer(_Dummy):
        @staticmethod
        def singleShot(*args, **kwargs):
            return None

    class QThread(_Dummy):
        def start(self):
            return None

    def Signal(*args, **kwargs):
        return _Dummy()

    # Core
    qtcore.Qt = _Qt
    qtcore.QTimer = QTimer
    qtcore.QThread = QThread
    qtcore.Signal = Signal

    # Gui
    qtgui.QColor = type('QColor', (_Dummy,), {})
    qtgui.QFont = type('QFont', (_Dummy,), {})

    # Widgets
    widget_names = [
        'QApplication', 'QCheckBox', 'QComboBox', 'QDialog', 'QFileDialog', 'QFrame',
        'QGridLayout', 'QHBoxLayout', 'QLabel', 'QLineEdit', 'QMainWindow', 'QMessageBox',
        'QProgressBar', 'QPushButton', 'QScrollArea', 'QTableWidget', 'QTableWidgetItem',
        'QTabWidget', 'QTextEdit', 'QVBoxLayout', 'QWidget', 'QHeaderView'
    ]
    for name in widget_names:
        setattr(qtwidgets, name, type(name, (_Dummy,), {}))

    pyside6.QtCore = qtcore
    pyside6.QtGui = qtgui
    pyside6.QtWidgets = qtwidgets

    sys.modules['PySide6'] = pyside6
    sys.modules['PySide6.QtCore'] = qtcore
    sys.modules['PySide6.QtGui'] = qtgui
    sys.modules['PySide6.QtWidgets'] = qtwidgets
