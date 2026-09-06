import os
from pathlib import Path
from string import Formatter

from aps_midi_prep_tool_app.bulk_extraction import (
    bulk_extract_images,
    discover_image_files,
    sanitize_output_folder_name,
)
from aps_midi_prep_tool_app.eseq_pianodir import (
    PIANODIR_DISK_METADATA_OFFSET,
    PIANODIR_DISK_METADATA_SIZE,
    PIANODIR_TARGET_FILE_SIZE,
    PianodirMetadata,
    build_pianodir_metadata_bytes,
)
from aps_midi_prep_tool_app.floppy_image import ImageEntry, ImageListing
from aps_midi_prep_tool_app.message_catalog import (
    BULK_EXTRACTION_MESSAGE_IDS,
    MESSAGES,
    SUPPORTED_LANGUAGES,
    tr,
    translate_text,
)


class FakeImageSession:
    def __init__(self, image_path, files):
        self.image_path = image_path
        self.files = files
        self.temp_directory = Path(image_path).parent / f"_{Path(image_path).name}_files"
        self.temp_directory.mkdir(exist_ok=True)
        self.cleaned = False

    def list_entries(self):
        entries = [
            ImageEntry(path=name, size=len(data), packed_size=len(data))
            for name, data in self.files.items()
        ]
        return ImageListing(entries=entries, free_space=0, cluster_size=1024)

    def extract_file(self, image_path):
        output_path = self.temp_directory / image_path.replace("/", "_")
        output_path.write_bytes(self.files[image_path])
        return os.fspath(output_path)

    def cleanup(self):
        self.cleaned = True


def _pianodir_bytes(album_title):
    data = bytearray(PIANODIR_TARGET_FILE_SIZE)
    metadata = build_pianodir_metadata_bytes(PianodirMetadata(disk_title=album_title))
    start = PIANODIR_DISK_METADATA_OFFSET
    data[start:start + PIANODIR_DISK_METADATA_SIZE] = metadata
    return bytes(data)


def _minimal_midi_bytes():
    track = b"\x00\xFF\x2F\x00"
    return (
        b"MThd\x00\x00\x00\x06\x00\x00\x00\x01\x01\xE0"
        + b"MTrk"
        + len(track).to_bytes(4, "big")
        + track
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


def test_discover_image_files_only_scans_selected_directory(tmp_path):
    (tmp_path / "B.HFE").write_bytes(b"")
    (tmp_path / "a.img").write_bytes(b"")
    (tmp_path / "notes.txt").write_text("not an image", encoding="utf-8")
    nested = tmp_path / "nested"
    nested.mkdir()
    (nested / "ignored.img").write_bytes(b"")

    assert [Path(path).name for path in discover_image_files(tmp_path)] == ["a.img", "B.HFE"]


def test_bulk_extraction_uses_album_names_and_falls_back_to_image_names(tmp_path):
    source = tmp_path / "images"
    output = tmp_path / "output"
    source.mkdir()
    (source / "one.img").write_bytes(b"image")
    (source / "two.img").write_bytes(b"image")
    (source / "three.img").write_bytes(b"image")

    files_by_image = {
        "one.img": {"PIANODIR.FIL": _pianodir_bytes("Shared Album"), "SONG.FIL": b"song one"},
        "two.img": {"PIANODIR.FIL": _pianodir_bytes("Shared Album"), "SONG.FIL": b"song two"},
        "three.img": {"README.TXT": b"fallback"},
    }

    def session_loader(image_path, **_kwargs):
        return FakeImageSession(image_path, files_by_image[Path(image_path).name])

    result = bulk_extract_images(
        source,
        output,
        use_album_names=True,
        session_loader=session_loader,
    )

    assert result.images_processed == 3
    assert (output / "Shared Album" / "SONG.FIL").read_bytes() == b"song one"
    assert (output / "Shared Album_2" / "SONG.FIL").read_bytes() == b"song two"
    assert (output / "three" / "README.TXT").read_bytes() == b"fallback"


def test_bulk_extraction_reads_pdisk_album_titles_with_pianodir_priority(tmp_path):
    source = tmp_path / "images"
    source.mkdir()
    for name in ["catalog", "pianodir", "invalid"]:
        (source / f"{name}.img").write_bytes(b"image")
    pdisk = (
        b"PDISK   MNG   \r\nP.PLAYER      \r\nVer1.01DMV0.53\r\n"
        + b"Evening Jazz".ljust(64, b" ") + b"\r\n"
    )
    files = {
        "catalog.img": {"pdisk.mng": pdisk, "FIRST.MID": _minimal_midi_bytes()},
        "pianodir.img": {"PDISK.MNG": pdisk, "PIANODIR.FIL": _pianodir_bytes("Existing title")},
        "invalid.img": {"PDISK.MNG": b"truncated", "FIRST.MID": _minimal_midi_bytes()},
    }
    result = bulk_extract_images(
        source, tmp_path / "output", use_album_names=True,
        session_loader=lambda path, **kwargs: FakeImageSession(path, files[Path(path).name]),
    )
    assert {Path(path).name for path in result.output_directories} == {
        "Evening Jazz", "Existing title", "invalid",
    }
    assert result.errors == ()


def test_bulk_extraction_reads_lf_only_catalog_titles_and_album_names(tmp_path):
    source = tmp_path / "images"
    source.mkdir()
    (source / "numeric.img").write_bytes(b"image")
    pdisk = (
        b"PDISK   MNG   \r\nP.PLAYER      \r\nVer1.01DMV0.54\r\n"
        + b"Catalog album".ljust(64, b" ") + b"\r\nS           \r\n"
    ).replace(b"\r\n", b"\n")
    psong = _psong_bytes([
        ("01.MID", "First catalog title"), ("02.MID", "Second catalog title"),
    ]).replace(b"\r\n", b"\n")
    files = {
        "PDISK.MNG": pdisk, "PSONG.MNG": psong,
        "01.MID": _minimal_midi_bytes(), "02.MID": _minimal_midi_bytes(),
    }
    result = bulk_extract_images(
        source, tmp_path / "output", use_album_names=True, long_midi_filenames=True,
        session_loader=lambda path, **kwargs: FakeImageSession(path, files),
    )
    assert result.errors == ()
    album = tmp_path / "output" / "Catalog album"
    assert sorted(path.name for path in album.glob("*.mid")) == [
        "01 - First catalog title.mid", "02 - Second catalog title.mid",
    ]
    assert (album / "PDISK.MNG").read_bytes() == pdisk
    assert (album / "PSONG.MNG").read_bytes() == psong


def test_bulk_extraction_conversion_omits_eseq_and_yamaha_directory_files(tmp_path):
    source = tmp_path / "images"
    output = tmp_path / "output"
    source.mkdir()
    (source / "disk.img").write_bytes(b"image")
    source_files = {
        "PIANODIR.FIL": _pianodir_bytes("Converted Album"),
        "MUSIC.DIR": b"clavinova directory metadata",
        "MUSIC/SONG.FIL": b"eseq source",
        "MUSIC/SONG.MID": b"existing midi",
        "NOTES.TXT": b"keep this",
    }

    def session_loader(image_path, **_kwargs):
        return FakeImageSession(image_path, source_files)

    def convert_to_midi(source_path, destination_path):
        Path(destination_path).write_bytes(b"converted " + Path(source_path).read_bytes())

    result = bulk_extract_images(
        source,
        output,
        convert_eseq=True,
        use_album_names=True,
        session_loader=session_loader,
        eseq_detector=lambda path: path.lower().endswith(".fil"),
        eseq_converter=convert_to_midi,
    )

    album_output = output / "Converted Album"
    music = album_output / "MUSIC"
    assert not (music / "SONG.FIL").exists()
    assert not (album_output / "PIANODIR.FIL").exists()
    assert not (album_output / "MUSIC.DIR").exists()
    assert (music / "SONG.MID").read_bytes() == b"existing midi"
    assert (music / "SONG_converted.mid").read_bytes() == b"converted eseq source"
    assert (album_output / "NOTES.TXT").read_bytes() == b"keep this"
    assert result.files_extracted == 2
    assert result.files_converted == 1
    assert result.included_eseq_sources is False
    assert result.errors == ()


def test_bulk_extraction_can_include_eseq_sources_with_midi_conversions(tmp_path):
    source = tmp_path / "images"
    output = tmp_path / "output"
    source.mkdir()
    (source / "disk.img").write_bytes(b"image")
    source_files = {
        "PIANODIR.FIL": _pianodir_bytes("Source Album"),
        "MUSIC.DIR": b"clavinova directory metadata",
        "MUSIC/SONG.FIL": b"eseq source",
        "NOTES.TXT": b"keep this",
    }

    def session_loader(image_path, **_kwargs):
        return FakeImageSession(image_path, source_files)

    def convert_to_midi(source_path, destination_path):
        Path(destination_path).write_bytes(b"converted " + Path(source_path).read_bytes())

    result = bulk_extract_images(
        source,
        output,
        convert_eseq=True,
        include_eseq_sources=True,
        use_album_names=True,
        session_loader=session_loader,
        eseq_detector=lambda path: path.lower().endswith(".fil"),
        eseq_converter=convert_to_midi,
    )

    album_output = output / "Source Album"
    assert (album_output / "PIANODIR.FIL").exists()
    assert (album_output / "MUSIC.DIR").read_bytes() == b"clavinova directory metadata"
    assert (album_output / "MUSIC" / "SONG.FIL").read_bytes() == b"eseq source"
    assert (album_output / "MUSIC" / "SONG.mid").read_bytes() == b"converted eseq source"
    assert (album_output / "NOTES.TXT").read_bytes() == b"keep this"
    assert result.files_extracted == 4
    assert result.files_converted == 1
    assert result.included_eseq_sources is True
    assert result.errors == ()


def test_bulk_extraction_can_use_long_filenames_and_trim_converted_titles(tmp_path):
    source = tmp_path / "images"
    output = tmp_path / "output"
    source.mkdir()
    (source / "disk.img").write_bytes(b"image")
    source_files = {
        "MUSIC/FIRST.FIL": b"  Moon   River  ",
        "MUSIC/SECOND.FIL": b" Summer   Wind ",
        "NOTES.TXT": b"keep this",
    }
    converted_titles = []

    def session_loader(image_path, **_kwargs):
        return FakeImageSession(image_path, source_files)

    def convert_to_midi(source_path, destination_path, *, title_override=None):
        converted_titles.append(title_override)
        Path(destination_path).write_bytes(str(title_override).encode("utf-8"))

    result = bulk_extract_images(
        source,
        output,
        convert_eseq=True,
        long_midi_filenames=True,
        trim_title_spaces=True,
        session_loader=session_loader,
        eseq_detector=lambda path: path.lower().endswith(".fil"),
        eseq_converter=convert_to_midi,
        eseq_title_reader=lambda path: Path(path).read_text(encoding="utf-8"),
    )

    music = output / "disk" / "MUSIC"
    assert (music / "01 - Moon River.mid").read_text(encoding="utf-8") == "Moon River"
    assert (music / "02 - Summer Wind.mid").read_text(encoding="utf-8") == "Summer Wind"
    assert converted_titles == ["Moon River", "Summer Wind"]
    assert (output / "disk" / "NOTES.TXT").read_bytes() == b"keep this"
    assert result.files_converted == 2
    assert result.errors == ()


def test_bulk_extraction_long_filenames_use_smart_pianosoft_catalog_for_existing_midi(
    tmp_path,
):
    source = tmp_path / "images"
    output = tmp_path / "output"
    source.mkdir()
    (source / "disk.img").write_bytes(b"image")
    midi_bytes = _minimal_midi_bytes()
    source_files = {
        "PSONG.MNG": _psong_bytes(
            [
                ("FIRST.MID", "  Moon   River  "),
                ("SECOND.MID", "  Summer   Wind  "),
            ]
        ),
        # Deliberately reverse filesystem order to verify catalog numbering.
        "SECOND.MID": midi_bytes,
        "FIRST.MID": midi_bytes,
    }

    def session_loader(image_path, **_kwargs):
        return FakeImageSession(image_path, source_files)

    result = bulk_extract_images(
        source,
        output,
        long_midi_filenames=True,
        session_loader=session_loader,
    )

    disk_output = output / "disk"
    assert (disk_output / "01 - Moon River.mid").read_bytes() == midi_bytes
    assert (disk_output / "02 - Summer Wind.mid").read_bytes() == midi_bytes
    assert not (disk_output / "FIRST.MID").exists()
    assert not (disk_output / "SECOND.MID").exists()
    assert (disk_output / "PSONG.MNG").exists()
    assert result.files_converted == 0
    assert result.files_extracted == 3
    assert result.errors == ()


def test_bulk_extraction_reports_stable_overall_and_per_image_progress(tmp_path):
    source = tmp_path / "images"
    output = tmp_path / "output"
    source.mkdir()
    (source / "one.img").write_bytes(b"image")
    (source / "two.img").write_bytes(b"image")
    files_by_image = {
        "one.img": {"ONE.MID": b"one", "TWO.MID": b"two"},
        "two.img": {"THREE.MID": b"three"},
    }
    details = []

    def session_loader(image_path, **_kwargs):
        return FakeImageSession(image_path, files_by_image[Path(image_path).name])

    bulk_extract_images(
        source,
        output,
        session_loader=session_loader,
        progress_detail_callback=details.append,
    )

    assert details[0]["stage"] == "scan"
    assert {detail["image_total"] for detail in details} == {2}
    assert [detail["overall_completed"] for detail in details] == sorted(
        detail["overall_completed"] for detail in details
    )
    assert [
        (detail["image_index"], detail["file_completed"], detail["file_total"])
        for detail in details
        if detail["stage"] == "extracting"
    ] == [(1, 0, 2), (1, 1, 2), (2, 0, 1)]
    assert details[-1]["stage"] == "finished"
    assert details[-1]["overall_completed"] == 2


def test_sanitize_output_folder_name_is_portable():
    assert sanitize_output_folder_name('  Album: One / Two  ') == "Album One Two"
    assert sanitize_output_folder_name("CON") == "CON Image"


def test_bulk_extraction_messages_cover_every_supported_language():
    supported_codes = {language.code for language in SUPPORTED_LANGUAGES}
    formatter = Formatter()
    for message_id in BULK_EXTRACTION_MESSAGE_IDS:
        assert set(MESSAGES[message_id]) == supported_codes
        assert all(MESSAGES[message_id][code].strip() for code in supported_codes)
        expected_fields = {
            field_name
            for _literal, field_name, _format_spec, _conversion in formatter.parse(
                MESSAGES[message_id]["en"]
            )
            if field_name
        }
        for language_code in supported_codes:
            translated_fields = {
                field_name
                for _literal, field_name, _format_spec, _conversion in formatter.parse(
                    MESSAGES[message_id][language_code]
                )
                if field_name
            }
            assert translated_fields == expected_fields


def test_bulk_extraction_action_and_progress_are_localized():
    assert translate_text("Bulk Extraction...", "bg") == "Масово извличане..."
    assert tr("bulk.action", "ja") == "一括抽出..."
    assert tr(
        "bulk.progress.extracting",
        "zh-Hans",
        image="disk.img",
        current=1,
        total=2,
        path="SONG.FIL",
    ) == "正在提取 disk.img：1/2（SONG.FIL）..."
