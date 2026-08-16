import re
from pathlib import Path
from string import Formatter

from aps_midi_prep_tool_app.app_info import APP_VERSION
from aps_midi_prep_tool_app.message_catalog import (
    COMMON_TEXT_TRANSLATIONS,
    MESSAGES,
    SUPPORTED_LANGUAGES,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _format_fields(text):
    return sorted(
        field_name
        for _, field_name, _, _ in Formatter().parse(text)
        if field_name is not None
    )


def test_release_version_is_consistent_across_app_and_documentation():
    readme = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")
    changelog = (PROJECT_ROOT / "CHANGELOG.md").read_text(encoding="utf-8")

    readme_match = re.search(
        r"^Current version: \x60([^\x60]+)\x60$",
        readme,
        re.MULTILINE,
    )
    changelog_match = re.search(
        r"^## \[([^]]+)] - \d{4}-\d{2}-\d{2}$",
        changelog,
        re.MULTILINE,
    )

    assert readme_match is not None
    assert changelog_match is not None
    assert readme_match.group(1) == APP_VERSION
    assert changelog_match.group(1) == APP_VERSION


def test_message_catalog_has_complete_language_and_placeholder_coverage():
    supported_codes = {language.code for language in SUPPORTED_LANGUAGES}

    for message_id, translations in MESSAGES.items():
        assert supported_codes <= set(translations), message_id
        expected_fields = _format_fields(translations["en"])
        for language_code in supported_codes:
            assert translations[language_code], (message_id, language_code)
            assert _format_fields(translations[language_code]) == expected_fields, (
                message_id,
                language_code,
            )


def test_common_text_catalog_has_complete_translation_and_placeholder_coverage():
    translated_codes = {
        language.code for language in SUPPORTED_LANGUAGES if language.code != "en"
    }

    for source, translations in COMMON_TEXT_TRANSLATIONS.items():
        assert translated_codes <= set(translations), source
        expected_fields = _format_fields(source)
        for language_code in translated_codes:
            assert translations[language_code], (source, language_code)
            assert _format_fields(translations[language_code]) == expected_fields, (
                source,
                language_code,
            )
