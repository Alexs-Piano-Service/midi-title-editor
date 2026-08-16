import inspect
import struct
from types import SimpleNamespace

import pytest

from aps_midi_prep_tool_app.main_window import MidiTitleWindow, _inspect_midi_bytes
from aps_midi_prep_tool_app.midi_type0_converter import (
    _encode_vlq,
    _parse_midi_chunks,
    _parse_track_events,
    analyze_pedal_softening_midi_bytes,
    apply_pedal_compatibility_to_midi_bytes,
)


def _track(events, end_tick=None):
    payload = bytearray()
    previous_tick = 0
    ordered_events = sorted(
        enumerate(events),
        key=lambda item: (item[1][0], item[0]),
    )
    for _source_order, (tick, raw) in ordered_events:
        payload.extend(_encode_vlq(tick - previous_tick))
        payload.extend(raw)
        previous_tick = tick
    final_tick = previous_tick if end_tick is None else end_tick
    payload.extend(_encode_vlq(final_tick - previous_tick))
    payload.extend(b"\xFF\x2F\x00")
    return b"MTrk" + len(payload).to_bytes(4, "big") + bytes(payload)


def _midi(*tracks, format_type=1, division=480, extra_chunks=()):
    return (
        struct.pack(">4sIHHH", b"MThd", 6, format_type, len(tracks), division)
        + b"".join(extra_chunks)
        + b"".join(tracks)
    )


def _raw_chunk(chunk_id, payload):
    return chunk_id + len(payload).to_bytes(4, "big") + payload


def _parsed_tracks(midi_bytes):
    _header_end, _format_type, _track_count, chunks = _parse_midi_chunks(midi_bytes)
    parsed = []
    for chunk in chunks:
        if chunk["id"] != b"MTrk":
            continue
        events, end_tick = _parse_track_events(
            midi_bytes[chunk["data_start"]:chunk["data_end"]]
        )
        parsed.append((events, end_tick))
    return parsed


def _sustain_events(events, channel=None):
    return [
        (tick, raw)
        for tick, _order, raw in events
        if len(raw) >= 3
        and (raw[0] & 0xF0) == 0xB0
        and raw[1] == 64
        and (channel is None or (raw[0] & 0x0F) == channel)
    ]


def _binary_pedal_track(channel=0):
    status = 0xB0 | channel
    return _track(
        [
            (0, b"\xFF\x51\x03\x07\xA1\x20"),
            (0, bytes([status, 64, 0])),
            (0, bytes([0x90 | channel, 60, 80])),
            (480, bytes([status, 64, 127])),
            (480, bytes([0x90 | channel, 64, 76])),
            (720, bytes([0x80 | channel, 60, 64])),
            (960, bytes([status, 64, 0])),
            (960, bytes([0x80 | channel, 64, 64])),
        ],
        end_tick=1200,
    )


def test_softens_binary_sustain_with_original_threshold_timing_and_preserves_other_events():
    source = _midi(
        _binary_pedal_track(),
        extra_chunks=(_raw_chunk(b"TEST", b"keep-me"),),
    )
    before_events = _parsed_tracks(source)[0][0]
    before_non_pedal = [
        (tick, raw)
        for tick, _order, raw in before_events
        if not (len(raw) >= 3 and (raw[0] & 0xF0) == 0xB0 and raw[1] == 64)
    ]

    analysis = analyze_pedal_softening_midi_bytes(source)
    assert analysis == {
        "classification": "binary",
        "stream_count": 1,
        "binary_stream_count": 1,
        "continuous_stream_count": 0,
        "static_stream_count": 0,
        "pedal_message_count": 3,
        "channels": [1],
    }

    softened, changed = apply_pedal_compatibility_to_midi_bytes(
        source,
        soften_sustain_pedal=True,
        pedal_down_ms=100,
        pedal_release_ms=180,
    )

    assert changed
    header_end, _format_type, _track_count, chunks = _parse_midi_chunks(softened)
    del header_end
    test_chunk = next(chunk for chunk in chunks if chunk["id"] == b"TEST")
    assert softened[test_chunk["data_start"]:test_chunk["data_end"]] == b"keep-me"

    after_events = _parsed_tracks(softened)[0][0]
    after_non_pedal = [
        (tick, raw)
        for tick, _order, raw in after_events
        if not (len(raw) >= 3 and (raw[0] & 0xF0) == 0xB0 and raw[1] == 64)
    ]
    assert after_non_pedal == before_non_pedal

    sustain = _sustain_events(after_events)
    assert (480, bytes([0xB0, 64, 64])) in sustain
    assert (960, bytes([0xB0, 64, 63])) in sustain
    assert any(0 < raw[2] < 127 for _tick, raw in sustain)
    assert all(raw[2] < 64 for tick, raw in sustain if tick < 480)
    assert all(raw[2] >= 64 for tick, raw in sustain if 480 < tick < 960)
    assert analyze_pedal_softening_midi_bytes(softened)["classification"] == "continuous"


def test_mixed_file_softens_only_binary_stream_and_preserves_continuous_stream():
    continuous_events = [
        (0, bytes([0xB1, 64, 0])),
        (300, bytes([0xB1, 64, 24])),
        (360, bytes([0xB1, 64, 55])),
        (420, bytes([0xB1, 64, 88])),
        (480, bytes([0xB1, 64, 127])),
        (900, bytes([0xB1, 64, 80])),
        (930, bytes([0xB1, 64, 40])),
        (960, bytes([0xB1, 64, 0])),
    ]
    source = _midi(
        _track(
            [
                (0, bytes([0xB0, 64, 0])),
                *continuous_events,
                (480, bytes([0xB0, 64, 127])),
                (960, bytes([0xB0, 64, 0])),
            ],
            end_tick=1000,
        )
    )
    assert analyze_pedal_softening_midi_bytes(source)["classification"] == "mixed"
    before_continuous = _sustain_events(_parsed_tracks(source)[0][0], channel=1)

    softened, changed = apply_pedal_compatibility_to_midi_bytes(
        source,
        soften_sustain_pedal=True,
        pedal_down_ms=55,
        pedal_release_ms=85,
    )

    assert changed
    output_events = _parsed_tracks(softened)[0][0]
    assert _sustain_events(output_events, channel=1) == before_continuous
    assert len(_sustain_events(output_events, channel=0)) > 3


def test_continuous_pedal_only_is_skipped_without_rewriting_file():
    source = _midi(
        _track(
            [
                (0, bytes([0xB0, 64, 0])),
                (120, bytes([0xB0, 64, 24])),
                (240, bytes([0xB0, 64, 58])),
                (360, bytes([0xB0, 64, 96])),
                (480, bytes([0xB0, 64, 127])),
                (600, bytes([0xB0, 64, 78])),
                (720, bytes([0xB0, 64, 32])),
                (840, bytes([0xB0, 64, 0])),
            ]
        )
    )

    output, changed = apply_pedal_compatibility_to_midi_bytes(
        source,
        soften_sustain_pedal=True,
    )

    assert not changed
    assert output is source
    assert analyze_pedal_softening_midi_bytes(source)["classification"] == "continuous"


def test_type_two_softening_uses_each_tracks_tempo_map():
    def pedal_track(tempo_bytes, channel):
        return _track(
            [
                (0, b"\xFF\x51\x03" + tempo_bytes),
                (0, bytes([0xB0 | channel, 64, 0])),
                (480, bytes([0xB0 | channel, 64, 127])),
                (960, bytes([0xB0 | channel, 64, 0])),
            ],
            end_tick=960,
        )

    source = _midi(
        pedal_track(b"\x07\xA1\x20", 0),
        pedal_track(b"\x0F\x42\x40", 1),
        format_type=2,
    )
    output, changed = apply_pedal_compatibility_to_midi_bytes(
        source,
        soften_sustain_pedal=True,
        pedal_down_ms=100,
        pedal_release_ms=180,
    )

    assert changed
    tracks = _parsed_tracks(output)
    first_pre_threshold = max(
        tick for tick, raw in _sustain_events(tracks[0][0], channel=0) if tick < 480 and raw[2] > 0
    )
    second_pre_threshold = max(
        tick for tick, raw in _sustain_events(tracks[1][0], channel=1) if tick < 480 and raw[2] > 0
    )
    assert first_pre_threshold < second_pre_threshold
    assert (480, bytes([0xB0, 64, 64])) in _sustain_events(tracks[0][0], channel=0)
    assert (480, bytes([0xB1, 64, 64])) in _sustain_events(tracks[1][0], channel=1)


def test_binary_conversion_and_softening_are_rejected_together():
    with pytest.raises(ValueError, match="cannot be combined"):
        apply_pedal_compatibility_to_midi_bytes(
            _midi(_binary_pedal_track()),
            binary_pedal=True,
            soften_sustain_pedal=True,
        )


def test_file_inspection_reports_binary_and_continuous_sustain_pedal_modes():
    binary_source = _midi(_binary_pedal_track())
    binary_inspection = _inspect_midi_bytes(
        binary_source,
        source_label="binary.mid",
    )
    assert binary_inspection["sustain_pedal_analysis"]["classification"] == "binary"
    assert (
        "Sustain pedal classification (CC64): Binary — on/off sustain data"
        in binary_inspection["metadata_text"]
    )

    continuous_source, changed = apply_pedal_compatibility_to_midi_bytes(
        binary_source,
        soften_sustain_pedal=True,
    )
    assert changed
    continuous_inspection = _inspect_midi_bytes(
        continuous_source,
        source_label="continuous.mid",
    )
    assert continuous_inspection["sustain_pedal_analysis"]["classification"] == "continuous"
    assert (
        "Sustain pedal classification (CC64): Continuous — graduated sustain data"
        in continuous_inspection["metadata_text"]
    )


def test_pedal_target_list_does_not_prefix_individual_song_names():
    source = inspect.getsource(
        MidiTitleWindow._pedal_compatibility_options_dialog
    )

    assert "One song:" not in source


def test_legacy_disklavier_option_is_described_as_a_channel_remap():
    source = inspect.getsource(
        MidiTitleWindow._pedal_compatibility_options_dialog
    )

    assert "Remap legacy Disklavier pedal (channel 3 → channel 1)" in source
    assert "Remap CC64, CC66, and CC67 from MIDI channel 3 to channel 1" in source
    assert "Repair legacy Disklavier" not in source


@pytest.mark.parametrize(
    ("target_index", "expected_rows"),
    [(-1, [(2, "first.mid"), (5, "second.mid")]), (1, [(5, "second.mid")])],
)
def test_pedal_utility_targets_all_rows_or_one_selected_song(target_index, expected_rows):
    rows = [(2, "first.mid"), (5, "second.mid")]
    captured = {}
    options = {
        "repair_disklavier_pedal": False,
        "binary_pedal": False,
        "pedal_cleanup": False,
        "virtual_piano_roll_pedal": False,
        "soften_sustain_pedal": True,
        "pedal_down_ms": 100,
        "pedal_release_ms": 180,
        "_target_index": target_index,
    }
    window = SimpleNamespace(
        choose_button=SimpleNamespace(isEnabled=lambda: True),
        _midi_rows_for_pedal_compatibility=lambda: rows,
        _pedal_compatibility_options_dialog=lambda _rows: dict(options),
        is_image_mode=lambda: False,
        _apply_pedal_compatibility_to_regular_rows=lambda selected, selected_options: captured.update(
            rows=selected,
            options=selected_options,
        ),
    )

    MidiTitleWindow.show_pedal_compatibility_utility(window)

    assert captured["rows"] == expected_rows
    assert captured["options"]["soften_sustain_pedal"] is True
    assert "_target_index" not in captured["options"]
