from types import SimpleNamespace
from unittest.mock import patch

import pytest
from PySide6.QtGui import QColor, QPalette

from aps_midi_prep_tool_app.main_window import (
    MidiTitleWindow,
    _build_dark_palette,
    _build_light_palette,
    _effective_appearance_mode,
)


def _relative_luminance(color):
    channels = []
    for value in (color.redF(), color.greenF(), color.blueF()):
        channels.append(value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def _contrast_ratio(first, second):
    light, dark = sorted(
        (_relative_luminance(QColor(first)), _relative_luminance(QColor(second))),
        reverse=True,
    )
    return (light + 0.05) / (dark + 0.05)


@pytest.mark.parametrize("palette_builder", [_build_light_palette, _build_dark_palette])
@pytest.mark.parametrize("group", [QPalette.Active, QPalette.Inactive, QPalette.Disabled])
def test_selected_text_meets_normal_text_contrast_target(palette_builder, group):
    palette = palette_builder()

    contrast = _contrast_ratio(
        palette.color(group, QPalette.Highlight),
        palette.color(group, QPalette.HighlightedText),
    )

    assert contrast >= 4.5


@pytest.mark.parametrize(
    ("system_palette", "expected_mode"),
    [(_build_light_palette(), "light"), (_build_dark_palette(), "dark")],
)
def test_system_theme_resolves_to_a_verified_palette(system_palette, expected_mode):
    assert _effective_appearance_mode("system", system_palette=system_palette) == expected_mode


@pytest.mark.parametrize("dark", [False, True])
def test_table_selection_style_covers_active_inactive_and_disabled_states(dark):
    table = SimpleNamespace(style_sheet=None)
    table.setStyleSheet = lambda value: setattr(table, "style_sheet", value)
    window = SimpleNamespace(table=table)

    with patch("aps_midi_prep_tool_app.main_window.is_dark_theme", return_value=dark):
        MidiTitleWindow._apply_table_selection_style(window)

    assert "QTableWidget::item:selected {" in table.style_sheet
    assert "QTableWidget::item:selected:!active {" in table.style_sheet
    assert "QTableWidget::item:selected:disabled {" in table.style_sheet
