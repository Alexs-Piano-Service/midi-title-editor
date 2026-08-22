import os
import struct

from aps_midi_prep_tool_app.midi_channel_merger import (
    merge_midi_channels_to_channel0_bytes,
    merge_midi_channels_to_channel0_path,
)
from aps_midi_prep_tool_app.midi_type0_converter import (
    _encode_vlq,
    _parse_track_events,
)


def _track(events, end_tick=None):
    payload = bytearray()
    previous_tick = 0
    for tick, raw in events:
        payload.extend(_encode_vlq(tick - previous_tick))
        payload.extend(raw)
        previous_tick = tick
    final_tick = previous_tick if end_tick is None else end_tick
    payload.extend(_encode_vlq(final_tick - previous_tick))
    payload.extend(b"\xFF\x2F\x00")
    return b"MTrk" + len(payload).to_bytes(4, "big") + bytes(payload)


def _midi(*chunks, format_type=0, division=480, header_extra=b"", trailing=b""):
    track_count = sum(chunk[:4] == b"MTrk" for chunk in chunks)
    header = struct.pack(">HHH", format_type, track_count, division) + header_extra
    return b"MThd" + len(header).to_bytes(4, "big") + header + b"".join(chunks) + trailing


def _declared_track_chunks(midi_bytes):
    header_end = 8 + int.from_bytes(midi_bytes[4:8], "big")
    track_count = int.from_bytes(midi_bytes[10:12], "big")
    tracks = []
    offset = header_end
    while len(tracks) < track_count:
        chunk_length = int.from_bytes(midi_bytes[offset + 4:offset + 8], "big")
        chunk_end = offset + 8 + chunk_length
        if midi_bytes[offset:offset + 4] == b"MTrk":
            tracks.append(midi_bytes[offset:chunk_end])
        offset = chunk_end
    return tracks, offset


def _events(track_chunk):
    return _parse_track_events(track_chunk[8:])


def _raw_events(track_chunk):
    events, _end_tick = _events(track_chunk)
    return [raw for _tick, _order, raw in events]


def test_type0_merges_all_seven_channel_voice_types_and_filters_commands():
    source = _midi(
        _track(
            [
                (0, b"\xFF\x20\x01\x09"),
                (0, b"\xFF\x20\x01\x10"),
                (0, b"\xC6\x2A"),
                (1, b"\x82\x3C\x40"),
                (2, b"\x93\x3E\x64"),
                (3, b"\xA4\x3E\x20"),
                (4, b"\xB5\x07\x63"),
                (5, b"\xD7\x31"),
                (6, b"\xE8\x00\x60"),
                (7, b"\xB9\x00\x01"),
                (8, b"\xB9\x20\x02"),
                (9, b"\xBA\x78\x00"),
                (10, b"\xBA\x7F\x00"),
            ],
            end_tick=20,
        )
    )

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert changed
    assert merged[8:14] == source[8:14]
    tracks, _trailing_start = _declared_track_chunks(merged)
    events, end_tick = _events(tracks[0])
    assert end_tick == 20
    raws = [raw for _tick, _order, raw in events]
    channel_events = [raw for raw in raws if 0x80 <= raw[0] <= 0xEF]
    assert all((raw[0] & 0x0F) == 0 for raw in channel_events)
    assert {raw[0] & 0xF0 for raw in channel_events} == {
        0x80,
        0x90,
        0xA0,
        0xB0,
        0xC0,
        0xD0,
        0xE0,
    }
    assert [raw for raw in channel_events if (raw[0] & 0xF0) == 0xC0] == [b"\xC0\x00"]
    assert [raw for raw in channel_events if (raw[0] & 0xF0) == 0xB0] == [b"\xB0\x07\x63"]
    assert b"\xFF\x20\x01\x00" in raws
    assert b"\xFF\x20\x01\x10" in raws


def test_type1_preserves_layout_metadata_sysex_unknown_chunks_and_trailing_data():
    conductor = _track(
        [
            (0, b"\xFF\x03\x09Conductor"),
            (0, b"\xFF\x51\x03\x07\xA1\x20"),
            (120, b"\xF0\x03\x43\x12\xF7"),
        ],
        end_tick=480,
    )
    first_notes = _track([(0, b"\x94\x3C\x64"), (240, b"\x84\x3C\x40")])
    second_notes = _track([(0, b"\x99\x24\x6E"), (240, b"\x89\x24\x40")])
    extra_chunk = b"JUNK" + (5).to_bytes(4, "big") + b"abcde"
    trailing = b"arbitrary trailing bytes that are not a chunk"
    source = _midi(
        conductor,
        extra_chunk,
        first_notes,
        second_notes,
        format_type=1,
        header_extra=b"\x12\x34",
        trailing=trailing,
    )

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert changed
    assert merged[:16] == source[:16]
    assert merged[8:10] == (1).to_bytes(2, "big")
    assert merged[10:12] == (3).to_bytes(2, "big")
    assert extra_chunk in merged
    assert merged.endswith(trailing)
    source_tracks, _source_trailing_start = _declared_track_chunks(source)
    merged_tracks, _merged_trailing_start = _declared_track_chunks(merged)
    assert merged_tracks[0] == source_tracks[0]
    assert sum(
        raw == b"\xC0\x00"
        for track in merged_tracks
        for raw in _raw_events(track)
    ) == 1
    assert b"\xC0\x00" in _raw_events(merged_tracks[1])
    assert b"\xC0\x00" not in _raw_events(merged_tracks[2])


def test_type2_injects_program_change_in_each_independent_note_track():
    source = _midi(
        _track([(0, b"\x91\x3C\x64"), (120, b"\x81\x3C\x40")]),
        _track([(0, b"\xC5\x28"), (10, b"\xE6\x00\x50")]),
        _track([(20, b"\x99\x24\x70"), (160, b"\x89\x24\x40")]),
        format_type=2,
    )

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert changed
    assert merged[8:10] == (2).to_bytes(2, "big")
    tracks, _trailing_start = _declared_track_chunks(merged)
    programs_by_track = [
        [raw for raw in _raw_events(track) if (raw[0] & 0xF0) == 0xC0]
        for track in tracks
    ]
    assert programs_by_track == [[b"\xC0\x00"], [], [b"\xC0\x00"]]
    assert b"\xE0\x00\x50" in _raw_events(tracks[1])


def test_smpte_division_is_preserved():
    smpte_division = 0xE728
    source = _midi(
        _track([(0, b"\x95\x40\x64"), (100, b"\x85\x40\x40")]),
        format_type=0,
        division=smpte_division,
    )

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert changed
    assert int.from_bytes(merged[12:14], "big") == smpte_division
    tracks, _trailing_start = _declared_track_chunks(merged)
    assert b"\x90\x40\x64" in _raw_events(tracks[0])


def test_running_status_channel_events_are_merged():
    payload = (
        b"\x00\x92\x3C\x64"
        b"\x0A\x3E\x6E"
        b"\x0A\x3C\x00"
        b"\x00\xFF\x2F\x00"
    )
    source = _midi(b"MTrk" + len(payload).to_bytes(4, "big") + payload)

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert changed
    tracks, _trailing_start = _declared_track_chunks(merged)
    events, end_tick = _events(tracks[0])
    assert end_tick == 20
    assert [(tick, raw) for tick, _order, raw in events] == [
        (0, b"\xC0\x00"),
        (0, b"\x90\x3C\x64"),
        (10, b"\x90\x3E\x6E"),
        (20, b"\x90\x3C\x00"),
    ]

    merged_again, changed_again = merge_midi_channels_to_channel0_bytes(merged)
    assert not changed_again
    assert merged_again is merged


def test_canonical_running_status_file_is_left_byte_for_byte_unchanged():
    payload = (
        b"\x00\xC0\x00"
        b"\x00\x90\x3C\x64"
        b"\x0A\x3E\x6E"
        b"\x0A\x3C\x00"
        b"\x00\xFF\x2F\x00"
    )
    source = _midi(b"MTrk" + len(payload).to_bytes(4, "big") + payload)

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert not changed
    assert merged is source


def test_midi_without_relevant_channel_or_prefix_events_is_an_exact_noop():
    source = _midi(
        _track(
            [
                (0, b"\xFF\x03\x04Meta"),
                (0, b"\xFF\x20\x01\x00"),
                (20, b"\xF7\x03\x01\x02\x03"),
            ],
            end_tick=100,
        )
    )

    merged, changed = merge_midi_channels_to_channel0_bytes(source)

    assert not changed
    assert merged is source


def test_path_api_replaces_atomically_and_skips_noop_destination(tmp_path, monkeypatch):
    source_path = tmp_path / "source.mid"
    destination_path = tmp_path / "merged.mid"
    source_path.write_bytes(
        _midi(_track([(0, b"\x94\x3C\x64"), (120, b"\x84\x3C\x40")]))
    )
    replace_calls = []
    real_replace = os.replace

    def recording_replace(source, destination):
        replace_calls.append((source, destination))
        real_replace(source, destination)

    monkeypatch.setattr(os, "replace", recording_replace)

    assert merge_midi_channels_to_channel0_path(source_path, destination_path)
    assert destination_path.exists()
    assert len(replace_calls) == 1
    assert os.path.dirname(replace_calls[0][0]) == os.path.dirname(
        os.fspath(destination_path)
    )
    assert replace_calls[0][1] == os.fspath(destination_path)
    assert not list(tmp_path.glob("*.aps_channel_merge_*.tmp"))

    noop_source = tmp_path / "noop.mid"
    noop_destination = tmp_path / "noop-result.mid"
    noop_source.write_bytes(_midi(_track([(0, b"\xFF\x03\x04Meta")])))
    assert not merge_midi_channels_to_channel0_path(noop_source, noop_destination)
    assert not noop_destination.exists()
