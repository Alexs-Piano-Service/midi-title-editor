import os
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QProgressDialog,
    QStyle,
    QStyleFactory,
    QWidget,
    QWidgetItem,
)

from aps_midi_prep_tool_app.main_window import (
    MidiTitleWindow,
    _TooltipDelayStyle,
    _build_dark_palette,
    _build_light_palette,
    _effective_appearance_mode,
)


def test_tooltip_style_sanitizes_internal_layout_items_before_base_call():
    app = QApplication.instance() or QApplication([])
    widget = QWidget()
    layout_item = QWidgetItem(widget)
    style = _TooltipDelayStyle(QStyleFactory.create("Fusion"))

    result = style.styleHint(
        QStyle.StyleHint.SH_ToolTip_WakeUpDelay,
        None,
        layout_item,
        None,
    )

    assert isinstance(result, int)
    widget.deleteLater()
    del app


def test_emulator_progress_dialog_keeps_a_stable_width_for_long_paths():
    app = QApplication.instance() or QApplication([])
    dialog = QProgressDialog("Preparing...", "Cancel", 0, 10)
    window = SimpleNamespace(
        _scaled_int=lambda value, minimum=0: max(int(minimum), int(value)),
    )
    window._set_progress_dialog_message = (
        lambda progress_dialog, message: MidiTitleWindow._set_progress_dialog_message(
            window,
            progress_dialog,
            message,
        )
    )

    MidiTitleWindow._stabilize_progress_dialog_width(window, dialog)
    long_message = (
        "Preparing /a/very/long/customer/library/path/"
        + "/nested-folder" * 40
        + "/Song.mid"
    )
    MidiTitleWindow._set_progress_dialog_message(window, dialog, long_message)

    assert dialog.minimumWidth() == 640
    assert dialog.maximumWidth() == 640
    assert dialog.labelText() != long_message
    assert dialog._aps_stable_progress_label.toolTip() == long_message
    dialog.deleteLater()
    del app


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
