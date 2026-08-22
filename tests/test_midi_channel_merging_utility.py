from types import SimpleNamespace

import pytest

from aps_midi_prep_tool_app.main_window import MidiTitleWindow


@pytest.mark.parametrize(
    ("target_index", "expected_rows"),
    [
        (-1, [(2, "first.mid"), (5, "second.mid")]),
        (1, [(5, "second.mid")]),
    ],
)
def test_channel_merge_utility_targets_all_rows_or_one_song(
    target_index,
    expected_rows,
):
    rows = [(2, "first.mid"), (5, "second.mid")]
    captured = {}
    window = SimpleNamespace(
        choose_button=SimpleNamespace(isEnabled=lambda: True),
        _midi_rows_for_channel_merging=lambda: rows,
        _channel_merging_options_dialog=lambda _rows: target_index,
        is_image_mode=lambda: False,
        _merge_channels_in_regular_rows=lambda selected: captured.update(rows=selected),
    )

    MidiTitleWindow.show_channel_merging_utility(window)

    assert captured["rows"] == expected_rows


@pytest.mark.parametrize(
    ("image_mode", "expected_handler"),
    [(False, "regular"), (True, "image")],
)
def test_channel_merge_utility_uses_the_current_staging_workflow(
    image_mode,
    expected_handler,
):
    rows = [(3, "song.mid")]
    captured = {}
    window = SimpleNamespace(
        choose_button=SimpleNamespace(isEnabled=lambda: True),
        _midi_rows_for_channel_merging=lambda: rows,
        _channel_merging_options_dialog=lambda _rows: -1,
        is_image_mode=lambda: image_mode,
        _merge_channels_in_regular_rows=lambda selected: captured.update(
            handler="regular",
            rows=selected,
        ),
        _merge_channels_in_image_rows=lambda selected: captured.update(
            handler="image",
            rows=selected,
        ),
    )

    MidiTitleWindow.show_channel_merging_utility(window)

    assert captured == {"handler": expected_handler, "rows": rows}
