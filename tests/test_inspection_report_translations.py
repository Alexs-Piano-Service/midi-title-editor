import struct

import pytest

from aps_midi_prep_tool_app.main_window import _inspect_midi_bytes
from aps_midi_prep_tool_app.message_catalog import SUPPORTED_LANGUAGES, translate_text
from aps_midi_prep_tool_app.midi_type0_converter import _encode_vlq


RAW_TITLE = "Save {filename} <song> & title"
RAW_FILENAME = "Save {count} <original> & song.mid"


def inspection_midi():
    title = RAW_TITLE.encode("utf-8")
    track = (
        b"\x00\xff\x03" + _encode_vlq(len(title)) + title
        + b"\x00\xc0\x00"  # Piano program on channel 1.
        + b"\x00\xb0\x07\x00"  # Channel 1 stays muted.
        + b"\x00\xb1\x0b\x00"  # Channel 2 expression later recovers.
        + b"\x00\xb0\x40\x7f"
        + b"\x00\x90\x3c\x64"
        + b"\x00\x91\x40\x64"
        + b"\x60\x80\x3c\x00"
        + b"\x00\x81\x40\x00"
        + b"\x00\xb1\x0b\x64"
        + b"\x00\xb0\x40\x00"
        + b"\x00\xff\x2f\x00"
    )
    return struct.pack(">4sIHHH", b"MThd", 6, 0, 1, 96) + b"MTrk" + len(track).to_bytes(4, "big") + track


@pytest.mark.parametrize("code", [language.code for language in SUPPORTED_LANGUAGES])
def test_inspection_report_localizes_labels_and_preserves_midi_data(code):
    data = inspection_midi()
    english = _inspect_midi_bytes(data, source_label=RAW_FILENAME)
    localized = _inspect_midi_bytes(data, source_label=RAW_FILENAME, language_code=code)
    report = localized["metadata_text"]

    assert RAW_FILENAME in report
    assert RAW_TITLE in report
    assert translate_text("Channel Summary:", code) in report
    assert translate_text("Track Name", code) in report
    assert translate_text("Pedals / Controllers:", code) in report
    assert translate_text("Other control changes:", code) in report
    assert translate_text("Mute / Volume Notes:", code) in report
    if code != "en":
        assert "Channel Summary:" not in report
        assert "Sustain pedal classification (CC64):" not in report

    for key in ("notes", "pedals", "duration", "channels", "track_count", "piano_channels", "sustain_pedal_analysis"):
        assert localized[key] == english[key]
    for channel in (1, 2):
        assert localized["channel_info"][channel]["mute_note"]
        assert localized["channel_info"][channel]["mute_note"] in report
        for key in ("note_count", "control_count", "programs", "piano_candidate"):
            assert localized["channel_info"][channel][key] == english["channel_info"][channel][key]


def test_default_inspection_report_keeps_english_export_wording():
    report = _inspect_midi_bytes(inspection_midi(), source_label=RAW_FILENAME)["metadata_text"]
    assert report.startswith(f"File: {RAW_FILENAME}\nMIDI type: Type 0\nTracks: 1 (declared 1)\n")
    assert "Channel 1: CC7 volume reaches 0; generic MIDI playback may mute that channel." in report
    assert "Channel 2: CC11 expression reaches 0 and later returns above 0." in report
    assert f"Track 1, tick 0: Track Name: {RAW_TITLE}" in report
