import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from .app_info import APP_VERSION, SETTINGS_APP, SETTINGS_ORG
from .onboarding_dialog import show_first_time_dialog
from .main_window import MidiTitleWindow


def _set_windows_app_id() -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            f"{SETTINGS_ORG}.{SETTINGS_APP}"
        )
    except Exception:
        pass


def main():
    _set_windows_app_id()
    app = QApplication(sys.argv)
    app.setOrganizationName(SETTINGS_ORG)
    app.setApplicationName(SETTINGS_APP)
    app.setApplicationVersion(APP_VERSION)
    app_icon = QIcon(str(Path(__file__).resolve().parent / "aps.ico"))
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    show_first_time_dialog()
    window = MidiTitleWindow()
    if not app_icon.isNull():
        window.setWindowIcon(app_icon)
    window.show()
    sys.exit(app.exec())
