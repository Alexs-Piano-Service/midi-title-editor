import ast
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QDialog, QDialogButtonBox, QWidget

from aps_midi_prep_tool_app import main_window
from aps_midi_prep_tool_app.message_catalog import (
    COMMON_TEXT_TRANSLATIONS,
    SUPPORTED_LANGUAGES,
    TEXT_TO_MESSAGE_ID,
    translate_text,
)


class DialogHarness(QWidget):
    _language_code = main_window.MidiTitleWindow._language_code
    _lt = main_window.MidiTitleWindow._lt
    _translate_dialog_tree = main_window.MidiTitleWindow._translate_dialog_tree
    _translate_dialog_button_box = main_window.MidiTitleWindow._translate_dialog_button_box

    def _has_pending_image_changes(self):
        return True


@pytest.mark.parametrize("code", [language.code for language in SUPPORTED_LANGUAGES])
def test_disk_confirmations_translate_before_inserting_user_data(monkeypatch, code):
    app = QApplication.instance() or QApplication([])
    window = DialogHarness()
    window.currentLanguage = code
    disk_format = SimpleNamespace(label="IBM 720K DD", size_bytes=737280)
    window.image_session = SimpleNamespace(disk_format=disk_format)
    target = "Save {original} <disk> & A:"
    captured = []

    def inspect_dialog(box):
        captured.append(box.text())
        assert target in box.text()
        assert box.defaultButton() is box.button(main_window.QMessageBox.No)
        assert box.button(main_window.QMessageBox.No).text() == translate_text("No", code)
        return main_window.QMessageBox.No

    monkeypatch.setattr(main_window.QMessageBox, "exec", inspect_dialog)
    methods = main_window.MidiTitleWindow
    assert not methods._confirm_format_floppy(window, target, disk_format)
    assert not methods._confirm_write_image_to_floppy(window, target, drive_size_bytes=1474560)
    assert not methods._confirm_save_to_floppy_files(window, target, drive_size_bytes=1474560)
    assert len(captured) == 3
    for text in captured:
        assert "{target}" not in text
        assert "{format}" not in text
        assert "{drive_size}" not in text
        if code != "en":
            assert "This will" not in text
            assert "The drive reports" not in text
    assert captured[1].count(translate_text("Pending image changes will be included.", code)) == 1
    window.close()
    app.processEvents()


@pytest.mark.parametrize("code", [language.code for language in SUPPORTED_LANGUAGES])
def test_dialog_translation_preserves_specific_action_buttons(code):
    app = QApplication.instance() or QApplication([])
    window = DialogHarness()
    window.currentLanguage = code
    dialog = QDialog(window)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
    buttons.button(QDialogButtonBox.Ok).setText(translate_text("Stage Conversion", code))

    window._translate_dialog_tree(dialog)
    window._translate_dialog_tree(dialog)

    assert buttons.button(QDialogButtonBox.Ok).text() == translate_text("Stage Conversion", code)
    assert buttons.button(QDialogButtonBox.Cancel).text() == translate_text("Cancel", code)
    window.close()
    app.processEvents()


def test_common_templates_format_values_without_translating_user_data():
    source = "A file named '{filename}' is already listed."
    for language in SUPPORTED_LANGUAGES:
        filename = "Save {count}.mid"
        assert filename in translate_text(source, language.code, filename=filename)
        assert "{filename}" not in translate_text(source, language.code, filename=filename)
        assert "{filename}" in translate_text(source, language.code)


@pytest.mark.parametrize("code", [language.code for language in SUPPORTED_LANGUAGES])
def test_message_details_toggle_stays_translated(code):
    app = QApplication.instance() or QApplication([])
    window = DialogHarness()
    window.currentLanguage = code
    box = main_window.QMessageBox(window)
    box.setDetailedText("Original diagnostic output")
    button = next(button for button in box.buttons() if button.property("_aps_details_translation"))
    assert button.text() == translate_text("Show Details...", code)
    button.click()
    assert button.text() == translate_text("Hide Details...", code)
    button.click()
    assert button.text() == translate_text("Show Details...", code)
    assert box.detailedText() == "Original diagnostic output"
    window.close()
    app.processEvents()


def test_dialog_error_and_confirmation_literals_have_translations():
    """Cover the implicit message-box path as well as explicit translation calls."""
    path = Path(main_window.__file__)
    tree = ast.parse(path.read_text(encoding="utf-8"))
    missing = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Attribute):
            continue
        name = node.func.attr
        if name in {"_show_operation_error", "_show_error_list"}:
            arguments = list(node.args[:2])
            arguments.extend(kw.value for kw in node.keywords if kw.arg == "guidance")
        elif name in {"information", "warning", "critical", "question"} and (
            isinstance(node.func.value, ast.Name) and node.func.value.id == "QMessageBox"
        ):
            arguments = node.args[1:3]
        else:
            continue
        for argument in arguments:
            if not isinstance(argument, ast.Constant) or not isinstance(argument.value, str):
                continue
            source = argument.value
            if not source:
                continue
            represented = source in COMMON_TEXT_TRANSLATIONS or source in TEXT_TO_MESSAGE_ID
            if source.endswith((":", "...")):
                represented = represented or source.rstrip(".:") in COMMON_TEXT_TRANSLATIONS
            if not represented:
                missing.append((node.lineno, source))
    assert not missing
