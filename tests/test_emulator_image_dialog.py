import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QWidget,
)

from aps_midi_prep_tool_app.main_window import MidiTitleWindow


class _Settings:
    def __init__(self, values=None):
        self.values = dict(values or {})

    def value(self, key, default=None, *, type=None):
        value = self.values.get(key, default)
        return type(value) if type is not None else value

    def setValue(self, key, value):
        self.values[key] = value


class _EmulatorDialogHarness(QWidget):
    SETTING_EMULATOR_IMAGE_SOURCE = MidiTitleWindow.SETTING_EMULATOR_IMAGE_SOURCE
    SETTING_EMULATOR_IMAGE_OUTPUT = MidiTitleWindow.SETTING_EMULATOR_IMAGE_OUTPUT
    SETTING_EMULATOR_IMAGE_PREFIX = MidiTitleWindow.SETTING_EMULATOR_IMAGE_PREFIX
    SETTING_EMULATOR_IMAGE_STARTING_NUMBER = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_STARTING_NUMBER
    )
    SETTING_EMULATOR_IMAGE_SAFETY_MARGIN_KIB = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_SAFETY_MARGIN_KIB
    )
    SETTING_EMULATOR_IMAGE_ALBUM_TITLE_OVERRIDE = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_ALBUM_TITLE_OVERRIDE
    )
    SETTING_EMULATOR_IMAGE_CONTENT = MidiTitleWindow.SETTING_EMULATOR_IMAGE_CONTENT
    SETTING_EMULATOR_IMAGE_OUTPUT_FORMAT = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_OUTPUT_FORMAT
    )
    SETTING_EMULATOR_IMAGE_DISK_FORMAT = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_DISK_FORMAT
    )
    SETTING_EMULATOR_IMAGE_INCLUDE_SUBFOLDERS = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_INCLUDE_SUBFOLDERS
    )
    SETTING_EMULATOR_IMAGE_SHUFFLE = MidiTitleWindow.SETTING_EMULATOR_IMAGE_SHUFFLE
    SETTING_EMULATOR_IMAGE_INCLUDE_SONG_LISTS = (
        MidiTitleWindow.SETTING_EMULATOR_IMAGE_INCLUDE_SONG_LISTS
    )

    show_emulator_image_utility = MidiTitleWindow.show_emulator_image_utility
    _language_code = MidiTitleWindow._language_code
    _t = MidiTitleWindow._t
    _lt = MidiTitleWindow._lt

    def __init__(self, source_directory):
        super().__init__()
        self.currentLanguage = "en"
        self.source_directory = os.fspath(source_directory)
        self.settings = _Settings(
            {self.SETTING_EMULATOR_IMAGE_INCLUDE_SUBFOLDERS: False}
        )
        self.recursive_option = None
        self.build_call = None

    def _disk_worker_busy(self):
        return False

    def _emulator_image_default_source_directory(self):
        return self.source_directory

    def _make_dialog_button_box(self, buttons, parent):
        return QDialogButtonBox(buttons, parent=parent)

    def _exec_child_dialog(self, dialog):
        checkbox = dialog.findChild(
            QCheckBox,
            "emulatorIncludeSubfoldersCheckbox",
        )
        self.recursive_option = {
            "text": checkbox.text(),
            "tooltip": checkbox.toolTip(),
            "checked": checkbox.isChecked(),
        }
        checkbox.setChecked(True)
        return QDialog.Accepted

    def _start_emulator_image_build(self, *args, **kwargs):
        self.build_call = (args, kwargs)


def test_emulator_disk_set_dialog_persists_and_forwards_recursive_option(tmp_path):
    app = QApplication.instance() or QApplication([])
    window = _EmulatorDialogHarness(tmp_path)

    window.show_emulator_image_utility()

    assert window.recursive_option == {
        "text": "Explore source folders recursively",
        "tooltip": (
            "Include supported MIDI and E-SEQ songs from every nested subfolder. "
            "Turn this off to use only songs directly in the selected source folder."
        ),
        "checked": False,
    }
    assert window.settings.values[
        window.SETTING_EMULATOR_IMAGE_INCLUDE_SUBFOLDERS
    ] is True
    assert window.build_call[1]["include_subfolders"] is True

    window.deleteLater()
    app.processEvents()
