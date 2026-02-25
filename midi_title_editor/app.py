import sys

from PySide6.QtWidgets import QApplication

from .onboarding_dialog import show_first_time_dialog
from .main_window import MidiTitleWindow


def main():
    app = QApplication(sys.argv)
    app.setOrganizationName("AlexPianoServiceLLC")
    app.setApplicationName("APSMidiTitleEditor")

    show_first_time_dialog()
    window = MidiTitleWindow()
    window.show()
    sys.exit(app.exec())
