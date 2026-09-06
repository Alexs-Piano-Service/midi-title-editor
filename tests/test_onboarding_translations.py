import html
import os
from string import Formatter

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QComboBox, QDialog, QLabel, QStackedWidget

from aps_midi_prep_tool_app import onboarding_dialog
from aps_midi_prep_tool_app.message_catalog import SUPPORTED_LANGUAGES, tr, translate_text
from aps_midi_prep_tool_app.onboarding_translations import ONBOARDING_TRANSLATIONS


def _fields(text):
    return sorted(field for _, field, _, _ in Formatter().parse(text) if field is not None)


def test_every_onboarding_language_has_all_pages_and_matching_placeholders():
    assert set(ONBOARDING_TRANSLATIONS) == {language.code for language in SUPPORTED_LANGUAGES}
    english = ONBOARDING_TRANSLATIONS["en"]
    page_ids = {message_id for _, message_id in onboarding_dialog.ONBOARDING_PAGES}
    assert set(english) == page_ids | {"welcome", "notice", "page_count", "workflow"}
    for code, messages in ONBOARDING_TRANSLATIONS.items():
        assert set(messages) == set(english), code
        for message_id, source in english.items():
            translated = messages[message_id]
            assert translated.strip(), (code, message_id)
            assert _fields(translated) == _fields(source), (code, message_id)
            if code != "en":
                assert translated != source, (code, message_id)


@pytest.mark.parametrize("code", [language.code for language in SUPPORTED_LANGUAGES])
def test_onboarding_displays_localized_bodies_notice_menu_paths_and_pagination(monkeypatch, code):
    app = QApplication.instance() or QApplication([])

    class Settings:
        def value(self, key, default=None, **_kwargs):
            return code if key == "language" else default

    monkeypatch.setattr(onboarding_dialog, "QSettings", lambda *_args: Settings())
    inspected = []

    def inspect_dialog(dialog):
        assert dialog.windowTitle() == ONBOARDING_TRANSLATIONS[code]["welcome"].format(
            app=onboarding_dialog.APP_TITLE_WITH_VERSION,
        )
        selector = dialog.findChild(QComboBox)
        stack = dialog.findChild(QStackedWidget)
        assert stack.count() == 11
        for index, (source_title, message_id) in enumerate(onboarding_dialog.ONBOARDING_PAGES):
            selector.setCurrentIndex(index)
            assert stack.currentIndex() == index
            labels = stack.currentWidget().findChildren(QLabel)
            assert labels[0].text() == translate_text(source_title, code)
            body = labels[1].text()
            assert html.escape(ONBOARDING_TRANSLATIONS[code]["notice"]) in body
            # Every prose fragment must come from the selected language. This
            # catches the old path that translated only titles and buttons.
            for literal, _, _, _ in Formatter().parse(ONBOARDING_TRANSLATIONS[code][message_id]):
                for paragraph_fragment in literal.split("\n\n"):
                    assert html.escape(paragraph_fragment) in body
            assert "{" not in body.split("<body>", 1)[1]
            count = ONBOARDING_TRANSLATIONS[code]["page_count"].format(current=index + 1, total=11)
            assert any(label.text() == count for label in dialog.findChildren(QLabel))
            if message_id == "builder":
                menu_path = " → ".join((translate_text("Utilities", code), tr("emulator.action", code)))
                assert html.escape(menu_path) in body
                assert html.escape(tr("emulator.include_song_lists", code)) in body
            if code != "en":
                assert ONBOARDING_TRANSLATIONS["en"]["notice"] not in body
        inspected.append(True)
        return QDialog.Accepted

    monkeypatch.setattr(QDialog, "exec", inspect_dialog)
    onboarding_dialog.show_first_time_dialog(force_show=True)
    assert inspected == [True]
    app.processEvents()
