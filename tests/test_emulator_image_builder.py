import hashlib
from pathlib import Path

import pytest

from aps_midi_prep_tool_app import emulator_image_builder
from aps_midi_prep_tool_app.emulator_image_builder import (
    build_emulator_disk_images,
    build_emulator_eseq_images,
    discover_midi_files,
    discover_song_files,
    sanitize_image_set_name,
)
from aps_midi_prep_tool_app.eseq_converter import (
    convert_midi_file_to_eseq_path,
    is_eseq_file,
)
from aps_midi_prep_tool_app.eseq_pianodir import (
    PIANODIR_COUNT_OFFSET,
    PIANODIR_FILENAME,
    read_pianodir_metadata_from_file,
)
from aps_midi_prep_tool_app.floppy_image import (
    DISK_FORMATS,
    FloppyImageError,
    FloppyImageSession,
    read_image_listing,
)
from aps_midi_prep_tool_app.midi_metadata import is_midi_file


def _midi_bytes(title):
    title_bytes = title.encode("ascii")
    track = (
        b"\x00\xff\x03"
        + bytes((len(title_bytes),))
        + title_bytes
        + b"\x00\x90\x3c\x40"
        + b"\x60\x80\x3c\x00"
        + b"\x00\xff\x2f\x00"
    )
    return (
        b"MThd"
        + (6).to_bytes(4, "big")
        + (0).to_bytes(2, "big")
        + (1).to_bytes(2, "big")
        + (96).to_bytes(2, "big")
        + b"MTrk"
        + len(track).to_bytes(4, "big")
        + track
    )


def _disk_format(key="ibm.720"):
    return next(item for item in DISK_FORMATS if item.key == key)


def test_discovers_midi_files_recursively_in_natural_order(tmp_path):
    source = tmp_path / "source"
    nested = source / "nested"
    nested.mkdir(parents=True)
    (source / "Song 10.mid").write_bytes(_midi_bytes("Ten"))
    (source / "Song 2.MIDI").write_bytes(_midi_bytes("Two"))
    (nested / "Song 3.mid").write_bytes(_midi_bytes("Three"))
    (nested / "ignore.txt").write_text("not MIDI", encoding="utf-8")
    (nested / "PIANODIR.FIL").write_bytes(b"directory metadata")

    discovered = discover_midi_files(source)

    assert [Path(path).name for path in discovered] == [
        "Song 2.MIDI",
        "Song 10.mid",
        "Song 3.mid",
    ]
    assert [Path(path).name for path in discover_midi_files(
        source,
        include_subfolders=False,
    )] == ["Song 2.MIDI", "Song 10.mid"]

    eseq_path = nested / "Song 4.fil"
    convert_midi_file_to_eseq_path(source / "Song 2.MIDI", eseq_path)
    assert [Path(path).name for path in discover_song_files(source)] == [
        "Song 2.MIDI",
        "Song 10.mid",
        "Song 3.mid",
        "Song 4.fil",
    ]


def test_sanitizes_portable_image_set_names():
    assert sanitize_image_set_name('  My: Emulator / Set?  ') == "My Emulator Set"
    assert sanitize_image_set_name("...") == "Emulator Disks"


def test_builds_multiple_verified_images_with_a_pianodir_per_disk(
    tmp_path,
    monkeypatch,
):
    source = tmp_path / "midi"
    output = tmp_path / "images"
    source.mkdir()
    for index in range(1, 4):
        (source / f"Song {index}.mid").write_bytes(
            _midi_bytes(f"Song {index}")
        )

    monkeypatch.setattr(emulator_image_builder, "PIANODIR_MAX_TRACKS", 2)
    result = build_emulator_eseq_images(
        source,
        output,
        set_name="Customer Set",
        album_title="Customer Album",
        catalog_number="APS-0001",
        disk_format=_disk_format(),
        output_ext="img",
    )

    assert result.midi_files_found == 3
    assert result.converted_files == 3
    assert result.images_created == 2
    assert [Path(path).name for path in result.output_paths] == [
        "Customer Set_01.img",
        "Customer Set_02.img",
    ]

    song_counts = []
    collected_names = []
    for image_path in result.output_paths:
        listing = read_image_listing(image_path)
        names = [entry.name for entry in listing.entries]
        assert names.count(PIANODIR_FILENAME) == 1
        eseq_names = [name for name in names if name.upper().endswith(".FIL") and name != PIANODIR_FILENAME]
        song_counts.append(len(eseq_names))
        collected_names.extend(eseq_names)

        session = FloppyImageSession.load(image_path)
        try:
            pianodir_path = session.extract_file(PIANODIR_FILENAME)
            pianodir_bytes = Path(pianodir_path).read_bytes()
            assert int.from_bytes(
                pianodir_bytes[PIANODIR_COUNT_OFFSET:PIANODIR_COUNT_OFFSET + 2],
                "little",
            ) == len(eseq_names) + 1
            metadata = read_pianodir_metadata_from_file(pianodir_path)
            assert metadata.catalog_number == "APS-0001"
            assert metadata.disk_title == "Customer Album"
            for eseq_name in eseq_names:
                assert is_eseq_file(session.extract_file(eseq_name))
        finally:
            session.cleanup()

    assert song_counts == [2, 1]
    assert len(set(collected_names)) == 3


def test_builds_midi_only_images_and_converts_eseq_only_when_needed(
    tmp_path,
    monkeypatch,
):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    midi_path = source / "Already MIDI.mid"
    midi_path.write_bytes(_midi_bytes("Already MIDI"))
    eseq_path = source / "Needs Conversion.fil"
    convert_midi_file_to_eseq_path(
        midi_path,
        eseq_path,
        filename_hint="NEEDSCON.FIL",
    )

    # The E-SEQ directory limit must not split a MIDI disk set.
    monkeypatch.setattr(emulator_image_builder, "PIANODIR_MAX_TRACKS", 1)
    result = build_emulator_disk_images(
        source,
        output,
        set_name="MIDI Set",
        disk_format=_disk_format(),
        output_content="midi",
    )

    assert result.song_files_found == 2
    assert result.files_prepared == 2
    assert result.converted_files == 1
    assert result.output_content == "midi"
    assert result.images_created == 1
    listing = read_image_listing(result.output_paths[0])
    names = [entry.name for entry in listing.entries]
    assert len(names) == 2
    assert PIANODIR_FILENAME not in names
    assert all(name.upper().endswith(".MID") for name in names)

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        assert all(is_midi_file(session.extract_file(name)) for name in names)
    finally:
        session.cleanup()


def test_builds_eseq_only_images_without_including_source_midi(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    midi_path = source / "Needs Conversion.mid"
    midi_path.write_bytes(_midi_bytes("Needs Conversion"))
    eseq_path = source / "Already ESEQ.fil"
    convert_midi_file_to_eseq_path(
        midi_path,
        eseq_path,
        filename_hint="ALREADYE.FIL",
    )

    result = build_emulator_disk_images(
        source,
        output,
        set_name="ESEQ Set",
        disk_format=_disk_format(),
        output_content="eseq",
    )

    assert result.files_prepared == 2
    assert result.converted_files == 1
    listing = read_image_listing(result.output_paths[0])
    names = [entry.name for entry in listing.entries]
    assert names.count(PIANODIR_FILENAME) == 1
    assert len(names) == 3
    assert not any(name.upper().endswith(".MID") for name in names)
    assert all(
        name == PIANODIR_FILENAME or name.upper().endswith(".FIL")
        for name in names
    )


def test_existing_output_is_not_overwritten(tmp_path):
    source = tmp_path / "midi"
    output = tmp_path / "images"
    source.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))
    output.mkdir()
    existing = output / "Existing.img"
    existing.write_bytes(b"keep this")
    before = hashlib.sha256(existing.read_bytes()).digest()

    with pytest.raises(FloppyImageError, match="already exists"):
        build_emulator_eseq_images(
            source,
            output,
            set_name="Existing",
            disk_format=_disk_format(),
        )

    assert hashlib.sha256(existing.read_bytes()).digest() == before
