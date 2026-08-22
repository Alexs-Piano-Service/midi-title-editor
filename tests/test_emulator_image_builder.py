import csv
import hashlib
from pathlib import Path

import pytest

from aps_midi_prep_tool_app import emulator_image_builder
from aps_midi_prep_tool_app.emulator_image_builder import (
    DEFAULT_SAFETY_MARGIN_BYTES,
    build_emulator_disk_images,
    build_emulator_eseq_images,
    discover_midi_files,
    discover_song_files,
    sanitize_image_prefix,
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
    FloppyOperationCancelled,
    read_image_listing,
)
from aps_midi_prep_tool_app.midi_metadata import (
    extract_eseq_title_from_file,
    extract_first_title_from_midi,
    is_midi_file,
)


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


def _vlq(value):
    encoded = [value & 0x7F]
    value >>= 7
    while value:
        encoded.append((value & 0x7F) | 0x80)
        value >>= 7
    return bytes(reversed(encoded))


def _large_midi_bytes(title, payload_size):
    title_bytes = title.encode("ascii")
    payload = b"x" * payload_size
    track = (
        b"\x00\xff\x03"
        + bytes((len(title_bytes),))
        + title_bytes
        + b"\x00\xff\x7f"
        + _vlq(len(payload))
        + payload
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


def _write_index_csv(directory, rows):
    fieldnames = (
        "number",
        "output_folder",
        "output_file",
        "title",
        "source_path",
        "sha256",
    )
    with (directory / "INDEX.csv").open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


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
    assert sanitize_image_prefix(" dsk-a ") == "DSKA"
    assert sanitize_image_prefix("library") == "LIBR"
    assert sanitize_image_prefix("---") == "DSKA"


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
        prefix="DSKB",
        starting_number=12,
        album_title="Customer Album",
        disk_format=_disk_format(),
        output_ext="img",
    )

    assert result.midi_files_found == 3
    assert result.converted_files == 3
    assert result.images_created == 2
    assert [Path(path).name for path in result.output_paths] == [
        "DSKB0012.img",
        "DSKB0013.img",
    ]
    assert result.image_prefix == "DSKB"
    assert result.starting_number == 12
    assert result.safety_margin_bytes == DEFAULT_SAFETY_MARGIN_BYTES

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
            disk_number = int(Path(image_path).stem[-4:])
            assert metadata.catalog_number == f"DSKB-{disk_number:04d}"
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
        output_ext="img",
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


def test_midi_images_preserve_titles_beyond_the_eseq_limit(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    long_title = "Mussorgsky - Bydlo (from Pictures at an Exhibition)"
    assert len(long_title) > 32
    source_path = source / "Bydlo.mid"
    source_path.write_bytes(_midi_bytes(long_title))

    result = build_emulator_disk_images(
        source,
        output,
        prefix="LONG",
        disk_format=_disk_format(),
        output_content="midi",
        output_ext="img",
        include_song_lists=True,
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        midi_entries = [
            entry
            for entry in session.list_entries().entries
            if entry.name.upper().endswith(".MID")
        ]
        assert len(midi_entries) == 1
        extracted_path = session.extract_file(midi_entries[0].path)
        assert Path(extracted_path).read_bytes() == source_path.read_bytes()
        assert extract_first_title_from_midi(extracted_path) == long_title
    finally:
        session.cleanup()

    assert long_title in Path(result.song_list_path).read_text(encoding="utf-8")


def test_index_csv_title_replaces_native_midi_title_without_truncation(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    source_path = source / "001 - Bydlo.mid"
    source_path.write_bytes(_midi_bytes("Embedded short title"))
    index_title = "Mussorgsky - Bydlo (from Pictures at an Exhibition)"
    _write_index_csv(
        source,
        [
            {
                "number": 1,
                "output_folder": ".",
                "output_file": source_path.name,
                "title": index_title,
            }
        ],
    )

    result = build_emulator_disk_images(
        source,
        output,
        prefix="INDX",
        disk_format=_disk_format(),
        output_content="midi",
        output_ext="img",
        include_song_lists=True,
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        midi_entry = next(
            entry
            for entry in session.list_entries().entries
            if entry.name.upper().endswith(".MID")
        )
        assert extract_first_title_from_midi(
            session.extract_file(midi_entry.path)
        ) == index_title
    finally:
        session.cleanup()

    assert index_title in Path(result.song_list_path).read_text(encoding="utf-8")


def test_index_csv_output_folder_disambiguates_duplicate_filenames(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    first_folder = source / "first"
    second_folder = source / "second"
    first_folder.mkdir(parents=True)
    second_folder.mkdir()
    (first_folder / "Song.mid").write_bytes(_midi_bytes("Embedded first"))
    (second_folder / "Song.mid").write_bytes(_midi_bytes("Embedded second"))
    first_title = "First indexed title"
    second_title = "Second indexed title"
    _write_index_csv(
        source,
        [
            {
                "number": 1,
                "output_folder": "first",
                "output_file": "Song.mid",
                "title": first_title,
            },
            {
                "number": 2,
                "output_folder": "second",
                "output_file": "Song.mid",
                "title": second_title,
            },
        ],
    )

    result = build_emulator_disk_images(
        source,
        output,
        prefix="PATH",
        disk_format=_disk_format(),
        output_content="midi",
        output_ext="img",
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        titles = {
            extract_first_title_from_midi(session.extract_file(entry.path))
            for entry in session.list_entries().entries
            if entry.name.upper().endswith(".MID")
        }
    finally:
        session.cleanup()

    assert titles == {first_title, second_title}


def test_index_csv_hash_match_restores_full_title_when_converting_eseq_to_midi(
    tmp_path,
):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    midi_path = source / "seed.mid"
    midi_path.write_bytes(_midi_bytes("Original title that reaches 32!!"))
    eseq_path = source / "LEGACY.FIL"
    convert_midi_file_to_eseq_path(midi_path, eseq_path)
    midi_path.unlink()
    index_title = "Mussorgsky - Ballet of the Unhatched Chickens (Pictures)"
    _write_index_csv(
        source,
        [
            {
                "number": 1,
                "output_folder": ".",
                "output_file": "A different filename.mid",
                "title": index_title,
                "source_path": "Archive/Another filename.fil",
                "sha256": hashlib.sha256(eseq_path.read_bytes()).hexdigest(),
            }
        ],
    )

    result = build_emulator_disk_images(
        source,
        output,
        prefix="HASH",
        disk_format=_disk_format(),
        output_content="midi",
        output_ext="img",
        include_song_lists=True,
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        midi_entry = next(
            entry
            for entry in session.list_entries().entries
            if entry.name.upper().endswith(".MID")
        )
        assert extract_first_title_from_midi(
            session.extract_file(midi_entry.path)
        ) == index_title
    finally:
        session.cleanup()

    assert index_title in Path(result.song_list_path).read_text(encoding="utf-8")


def test_eseq_images_and_song_lists_use_the_on_disk_32_byte_title(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    source_path = source / "Bydlo.mid"
    source_path.write_bytes(_midi_bytes("Embedded short title"))
    index_title = "Mussorgsky - Bydlo (from Pictures at an Exhibition)"
    expected_title = index_title[:32]
    _write_index_csv(
        source,
        [
            {
                "number": 1,
                "output_folder": ".",
                "output_file": source_path.name,
                "title": index_title,
            }
        ],
    )

    result = build_emulator_disk_images(
        source,
        output,
        prefix="ESEQ",
        disk_format=_disk_format(),
        output_content="eseq",
        output_ext="img",
        include_song_lists=True,
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        eseq_entries = [
            entry
            for entry in session.list_entries().entries
            if entry.name.upper().endswith(".FIL")
            and entry.name.upper() != PIANODIR_FILENAME
        ]
        assert len(eseq_entries) == 1
        extracted_path = session.extract_file(eseq_entries[0].path)
        assert extract_eseq_title_from_file(extracted_path) == expected_title
    finally:
        session.cleanup()

    song_list = Path(result.song_list_path).read_text(encoding="utf-8")
    assert f"1. {expected_title}\n" in song_list
    assert index_title not in song_list


def test_index_csv_updates_an_existing_eseq_title_with_the_32_byte_limit(
    tmp_path,
):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    seed_path = tmp_path / "seed.mid"
    seed_path.write_bytes(_midi_bytes("Embedded E-SEQ title"))
    eseq_path = source / "LEGACY.FIL"
    convert_midi_file_to_eseq_path(seed_path, eseq_path)
    index_title = "Rubinstein - Romance in E flat major, complete title"
    expected_title = index_title[:32]
    _write_index_csv(
        source,
        [
            {
                "number": 1,
                "output_folder": ".",
                "output_file": eseq_path.name,
                "title": index_title,
            }
        ],
    )

    result = build_emulator_disk_images(
        source,
        output,
        prefix="COPY",
        disk_format=_disk_format(),
        output_content="eseq",
        output_ext="img",
        include_song_lists=True,
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        eseq_entry = next(
            entry
            for entry in session.list_entries().entries
            if entry.name.upper().endswith(".FIL")
            and entry.name.upper() != PIANODIR_FILENAME
        )
        assert extract_eseq_title_from_file(
            session.extract_file(eseq_entry.path)
        ) == expected_title
    finally:
        session.cleanup()

    song_list = Path(result.song_list_path).read_text(encoding="utf-8")
    assert f"1. {expected_title}\n" in song_list
    assert index_title not in song_list


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
        output_ext="img",
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
    existing = output / "EXIS0001.img"
    existing.write_bytes(b"keep this")
    before = hashlib.sha256(existing.read_bytes()).digest()

    with pytest.raises(FloppyImageError, match="already exists"):
        build_emulator_eseq_images(
            source,
            output,
            set_name="Existing",
            disk_format=_disk_format(),
            output_ext="img",
        )

    assert hashlib.sha256(existing.read_bytes()).digest() == before


def test_existing_output_can_be_approved_for_replacement(tmp_path):
    source = tmp_path / "midi"
    output = tmp_path / "images"
    source.mkdir()
    output.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))
    existing = output / "REPL0007.img"
    existing.write_bytes(b"previous image")
    requested_paths = []

    result = build_emulator_disk_images(
        source,
        output,
        prefix="REPL",
        starting_number=7,
        disk_format=_disk_format(),
        output_ext="img",
        output_content="midi",
        overwrite_callback=lambda paths: requested_paths.extend(paths) or True,
    )

    assert requested_paths == [str(existing)]
    assert result.output_paths == (str(existing),)
    assert existing.read_bytes() != b"previous image"
    assert len(read_image_listing(existing).entries) == 1


def test_declining_output_replacement_preserves_existing_file(tmp_path):
    source = tmp_path / "midi"
    output = tmp_path / "images"
    source.mkdir()
    output.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))
    existing = output / "STAY0003.img"
    existing.write_bytes(b"keep this image")

    with pytest.raises(FloppyOperationCancelled, match="replacement was cancelled"):
        build_emulator_disk_images(
            source,
            output,
            prefix="STAY",
            starting_number=3,
            disk_format=_disk_format(),
            output_ext="img",
            output_content="midi",
            overwrite_callback=lambda _paths: False,
        )

    assert existing.read_bytes() == b"keep this image"


def test_replacement_commit_failure_restores_every_existing_image(
    tmp_path,
    monkeypatch,
):
    source = tmp_path / "midi"
    output = tmp_path / "images"
    source.mkdir()
    output.mkdir()
    for index in range(1, 3):
        (source / f"Song {index}.mid").write_bytes(
            _midi_bytes(f"Song {index}")
        )
    first = output / "ROLL0001.img"
    second = output / "ROLL0002.img"
    first.write_bytes(b"first existing image")
    second.write_bytes(b"second existing image")
    real_finish = emulator_image_builder._finish_temp_output
    commit_calls = []

    def fail_second_commit(staged_path, final_path):
        commit_calls.append(final_path)
        if len(commit_calls) == 2:
            raise OSError("simulated commit failure")
        return real_finish(staged_path, final_path)

    monkeypatch.setattr(emulator_image_builder, "PIANODIR_MAX_TRACKS", 1)
    monkeypatch.setattr(
        emulator_image_builder,
        "_finish_temp_output",
        fail_second_commit,
    )

    with pytest.raises(OSError, match="simulated commit failure"):
        build_emulator_disk_images(
            source,
            output,
            prefix="ROLL",
            starting_number=1,
            disk_format=_disk_format(),
            output_ext="img",
            output_content="eseq",
            overwrite_existing=True,
        )

    assert commit_calls == [str(first), str(second)]
    assert first.read_bytes() == b"first existing image"
    assert second.read_bytes() == b"second existing image"


def test_default_prefix_numbers_even_a_single_image(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))

    result = build_emulator_disk_images(
        source,
        output,
        disk_format=_disk_format(),
        output_ext="img",
        output_content="midi",
    )

    assert [Path(path).name for path in result.output_paths] == ["DSKA0001.img"]
    assert result.song_list_path == ""
    assert not list(output.glob("*.txt"))


def test_include_song_lists_writes_every_image_and_song_in_packed_order(
    tmp_path,
    monkeypatch,
):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    for index, title in enumerate(("First   Song", "Second Song", "Third Song"), start=1):
        (source / f"Song {index}.mid").write_bytes(_midi_bytes(title))

    monkeypatch.setattr(emulator_image_builder, "PIANODIR_MAX_TRACKS", 2)
    result = build_emulator_disk_images(
        source,
        output,
        prefix="LIST",
        starting_number=5,
        album_title="Customer Album",
        disk_format=_disk_format(),
        output_ext="img",
        output_content="eseq",
        include_song_lists=True,
    )

    assert Path(result.song_list_path).name == "LIST0005-LIST0006-song-lists.txt"
    assert Path(result.song_list_path).read_text(encoding="utf-8") == (
        "Emulator Disk Set Song Lists\n"
        "Images: 2\n"
        "Songs: 3\n"
        "\n"
        "Image 1 of 2: LIST0005.img\n"
        "Album: Customer Album\n"
        "Catalog: LIST-0005\n"
        "\n"
        "1. First Song\n"
        "2. Second Song\n"
        "\n"
        "Image 2 of 2: LIST0006.img\n"
        "Album: Customer Album\n"
        "Catalog: LIST-0006\n"
        "\n"
        "1. Third Song\n"
    )


def test_existing_song_list_is_not_overwritten(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    output.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))
    existing = output / "KEEP0007-song-list.txt"
    existing.write_text("keep this", encoding="utf-8")

    with pytest.raises(FloppyImageError, match="already exists"):
        build_emulator_disk_images(
            source,
            output,
            prefix="KEEP",
            starting_number=7,
            disk_format=_disk_format(),
            output_ext="img",
            output_content="midi",
            include_song_lists=True,
        )

    assert existing.read_text(encoding="utf-8") == "keep this"
    assert not (output / "KEEP0007.img").exists()


def test_eseq_album_title_defaults_to_the_generated_catalog_id(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))

    result = build_emulator_disk_images(
        source,
        output,
        prefix="DSKC",
        starting_number=42,
        disk_format=_disk_format(),
        output_ext="img",
        output_content="eseq",
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        metadata = read_pianodir_metadata_from_file(
            session.extract_file(PIANODIR_FILENAME)
        )
    finally:
        session.cleanup()

    assert metadata.catalog_number == "DSKC-0042"
    assert metadata.disk_title == "DSKC-0042"


def test_shuffle_randomizes_the_order_before_disk_packing(tmp_path, monkeypatch):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    for index in range(1, 4):
        (source / f"Song {index}.mid").write_bytes(_midi_bytes(f"Song {index}"))

    monkeypatch.setattr(
        emulator_image_builder.random,
        "shuffle",
        lambda paths: paths.reverse(),
    )
    result = build_emulator_disk_images(
        source,
        output,
        shuffle=True,
        disk_format=_disk_format(),
        output_ext="img",
        output_content="midi",
        include_song_lists=True,
    )

    session = FloppyImageSession.load(result.output_paths[0])
    try:
        titles = [
            extract_first_title_from_midi(session.extract_file(entry.path))
            for entry in session.list_entries().entries
        ]
    finally:
        session.cleanup()

    assert result.shuffled is True
    assert titles == ["Song 3", "Song 2", "Song 1"]
    song_list_text = Path(result.song_list_path).read_text(encoding="utf-8")
    assert song_list_text.index("1. Song 3") < song_list_text.index("2. Song 2")
    assert song_list_text.index("2. Song 2") < song_list_text.index("3. Song 1")


def test_reserves_requested_free_space_when_packing_disks(tmp_path):
    source = tmp_path / "songs"
    output = tmp_path / "images"
    source.mkdir()
    for index in range(1, 3):
        (source / f"Large {index}.mid").write_bytes(
            _large_midi_bytes(f"Large {index}", 190_000)
        )

    margin = 400 * 1024
    result = build_emulator_disk_images(
        source,
        output,
        prefix="SAFE",
        starting_number=7,
        safety_margin_bytes=margin,
        disk_format=_disk_format(),
        output_ext="img",
        output_content="midi",
    )

    assert [Path(path).name for path in result.output_paths] == [
        "SAFE0007.img",
        "SAFE0008.img",
    ]
    assert all(
        read_image_listing(path).free_space >= margin
        for path in result.output_paths
    )


@pytest.mark.parametrize("starting_number", [-1, 10_000, "invalid"])
def test_rejects_invalid_starting_numbers(tmp_path, starting_number):
    source = tmp_path / "songs"
    source.mkdir()
    (source / "Song.mid").write_bytes(_midi_bytes("Song"))

    with pytest.raises(FloppyImageError, match="starting disk number"):
        build_emulator_disk_images(
            source,
            tmp_path / "images",
            starting_number=starting_number,
            output_ext="img",
            output_content="midi",
        )


def test_rejects_a_set_that_would_run_past_slot_9999(tmp_path, monkeypatch):
    source = tmp_path / "songs"
    source.mkdir()
    for index in range(2):
        (source / f"Song {index}.mid").write_bytes(_midi_bytes(f"Song {index}"))
    monkeypatch.setattr(emulator_image_builder, "PIANODIR_MAX_TRACKS", 1)

    with pytest.raises(FloppyImageError, match="exceed disk 9999"):
        build_emulator_disk_images(
            source,
            tmp_path / "images",
            starting_number=9999,
            output_ext="img",
            output_content="eseq",
        )
