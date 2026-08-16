from pathlib import Path
from types import MethodType, SimpleNamespace

from aps_midi_prep_tool_app.main_window import MidiTitleWindow
from aps_midi_prep_tool_app.smart_pianosoft import (
    parse_smart_pianosoft_song_catalog,
    update_smart_pianosoft_song_catalog,
)


def _psong_bytes(records):
    record_size = 0xB0
    data = bytearray(b" " * (0x80 + len(records) * record_size))
    data[0x00:0x10] = b"PSONG   MNG   \r\n"
    data[0x10:0x20] = f"MAX{len(records):03d}        \r\n".encode("ascii")
    data[0x20:0x30] = f"FILE{len(records):03d}       \r\n".encode("ascii")
    for index, (filename, title) in enumerate(records):
        start = 0x80 + index * record_size
        stem, extension = filename.rsplit(".", 1)
        data[start:start + 8] = stem.encode("cp1252").ljust(8, b" ")
        data[start + 8:start + 11] = extension.encode("cp1252").ljust(3, b" ")
        data[start + 14:start + 16] = b"\r\n"
        data[start + 16:start + 48] = title.encode("cp1252").ljust(32, b" ")
    return bytes(data)


def test_smart_pianosoft_catalog_reads_titles_and_track_order():
    songs = parse_smart_pianosoft_song_catalog(
        _psong_bytes(
            [
                ("FIRST.MID", "  Moon   River  "),
                ("SECOND.MID", "You Can't AlwaysGet What You..."),
            ]
        )
    )

    assert [(song.track_number, song.filename, song.title) for song in songs] == [
        (1, "FIRST.MID", "Moon River"),
        (2, "SECOND.MID", "You Can't Always Get What You..."),
    ]


def test_smart_pianosoft_catalog_title_update_preserves_other_records():
    original = _psong_bytes(
        [
            ("FIRST.MID", "Moon River"),
            ("SECOND.MID", "Summer Wind"),
        ]
    )

    patched = update_smart_pianosoft_song_catalog(
        original,
        {"FIRST.MID": "The New Title"},
    )
    songs = parse_smart_pianosoft_song_catalog(patched)

    assert [song.title for song in songs] == ["The New Title", "Summer Wind"]
    assert patched[:0x90] == original[:0x90]
    assert patched[0xB0:] == original[0xB0:]


def test_catalog_backed_image_title_edit_stages_psong_not_midi(tmp_path):
    catalog_path = tmp_path / "PSONG.MNG"
    catalog_path.write_bytes(_psong_bytes([("FIRST.MID", "Moon River")]))
    patched_dir = tmp_path / "patched"
    patched_dir.mkdir()
    info = {"title_source": "smart_pianosoft_catalog", "title": "Moon River"}
    window = SimpleNamespace(
        image_session=SimpleNamespace(
            patched_dir=str(patched_dir),
            extract_file=lambda _path: str(catalog_path),
        ),
        smartPianoSoftCatalogPath="PSONG.MNG",
        pendingSmartPianoSoftCatalogReplacement="",
        pendingSmartPianoSoftTitleEdits={},
        pendingImageReplacements={},
        pendingImageTitleEdits={"FIRST.MID": "must not reach the MIDI"},
        _image_title_is_smart_pianosoft_catalog_backed=lambda _path: True,
        _image_info_for_path=lambda _path: info,
    )
    window._stage_smart_pianosoft_catalog_title = MethodType(
        MidiTitleWindow._stage_smart_pianosoft_catalog_title,
        window,
    )

    destination = MidiTitleWindow._stage_image_title_edit(
        window,
        "FIRST.MID",
        "The New Title",
    )

    assert destination == "smart_pianosoft_catalog"
    assert "FIRST.MID" not in window.pendingImageTitleEdits
    replacement = window.pendingImageReplacements["PSONG.MNG"]
    assert parse_smart_pianosoft_song_catalog(Path(replacement).read_bytes())[0].title == (
        "The New Title"
    )
    assert parse_smart_pianosoft_song_catalog(catalog_path.read_bytes())[0].title == "Moon River"
    assert info["title"] == "The New Title"
