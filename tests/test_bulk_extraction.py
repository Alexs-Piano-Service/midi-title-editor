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
