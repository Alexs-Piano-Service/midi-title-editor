"""Build emulator-ready MIDI or Yamaha E-SEQ floppy-image sets."""

from __future__ import annotations

import os
import random
import re
import shutil
import tempfile
import uuid
from dataclasses import dataclass

from .dos83_renamer import build_dos83_filename
from .eseq_converter import (
    convert_eseq_file_to_midi_path,
    convert_midi_file_to_eseq_path,
    is_eseq_file,
)
from .eseq_pianodir import (
    PIANODIR_FILENAME,
    PIANODIR_MAX_TRACKS,
    PianodirMetadata,
    PianodirTrackEntry,
    build_eseq_order_key_from_path,
    build_pianodir_bytes,
    is_clavinova_mda_file,
    update_eseq_order_key_to_path,
)
from .floppy_image import (
    DISK_FORMATS,
    FloppyImageError,
    FloppyOperationCancelled,
    _copy_host_file_into_image,
    _delete_eseq_directory_entries_from_image,
    _finish_temp_output,
    _is_image_capacity_error,
    _write_image_direct,
    allocated_size,
    create_blank_floppy_image,
    read_image_listing,
)
from .midi_metadata import (
    extract_eseq_title_from_file,
    extract_first_title_from_midi,
    is_midi_file,
)
from .message_catalog import tr


MIDI_EXTENSIONS = {".mid", ".midi"}
ESEQ_EXTENSIONS = {".fil", ".mda"}
SONG_EXTENSIONS = MIDI_EXTENSIONS | ESEQ_EXTENSIONS
EMULATOR_IMAGE_EXTENSIONS = {"img", "hfe"}
EMULATOR_CONTENT_FORMATS = {"eseq", "midi"}
DEFAULT_IMAGE_PREFIX = "DSKA"
DEFAULT_STARTING_NUMBER = 1
DEFAULT_SAFETY_MARGIN_BYTES = 32 * 1024
MAX_IMAGE_NUMBER = 9999
_INVALID_PORTABLE_FILENAME_CHARS = '<>:"/\\|?*'
_NATURAL_NUMBER_RE = re.compile(r"(\d+)")


@dataclass(frozen=True)
class EmulatorImageBuildResult:
    source_directory: str
    output_directory: str
    song_files_found: int
    files_prepared: int
    converted_files: int
    images_created: int
    output_content: str
    output_paths: tuple[str, ...]
    image_prefix: str
    starting_number: int
    safety_margin_bytes: int
    shuffled: bool
    song_list_path: str = ""

    @property
    def midi_files_found(self):
        """Compatibility alias for callers of the original E-SEQ-only builder."""
        return self.song_files_found


@dataclass(frozen=True)
class _PreparedSong:
    source_path: str
    image_path: str
    local_path: str
    title: str


def _natural_sort_key(path):
    relative = os.fspath(path).replace("\\", "/")
    return tuple(
        int(part) if part.isdigit() else part.casefold()
        for part in _NATURAL_NUMBER_RE.split(relative)
    )


def discover_song_files(source_directory, *, include_subfolders=True):
    """Return candidate MIDI and E-SEQ song files in stable natural order."""
    source_directory = os.path.abspath(os.fspath(source_directory))
    if not os.path.isdir(source_directory):
        raise FloppyImageError(f"The MIDI folder was not found: {source_directory}")

    paths = []
    try:
        if include_subfolders:
            for root, directory_names, filenames in os.walk(source_directory):
                directory_names.sort(key=_natural_sort_key)
                for filename in filenames:
                    if (
                        os.path.splitext(filename)[1].lower() in SONG_EXTENSIONS
                        and filename.upper() not in {"PIANODIR.FIL", "MUSIC.DIR"}
                    ):
                        paths.append(os.path.abspath(os.path.join(root, filename)))
        else:
            with os.scandir(source_directory) as entries:
                for entry in entries:
                    if (
                        entry.is_file()
                        and os.path.splitext(entry.name)[1].lower() in SONG_EXTENSIONS
                        and entry.name.upper() not in {"PIANODIR.FIL", "MUSIC.DIR"}
                    ):
                        paths.append(os.path.abspath(entry.path))
    except OSError as exc:
        raise FloppyImageError(f"Could not read the MIDI folder: {exc}") from exc

    return sorted(
        paths,
        key=lambda path: (
            os.path.relpath(path, source_directory).count(os.sep),
            _natural_sort_key(os.path.relpath(path, source_directory)),
        ),
    )


def discover_midi_files(source_directory, *, include_subfolders=True):
    """Return only MIDI files; retained for API compatibility."""
    return [
        path
        for path in discover_song_files(
            source_directory,
            include_subfolders=include_subfolders,
        )
        if os.path.splitext(path)[1].lower() in MIDI_EXTENSIONS
    ]


def sanitize_image_set_name(name, fallback="Emulator Disks"):
    """Return a portable filename stem for the generated image set."""
    text = re.sub(r"\s+", " ", str(name or "")).strip()
    cleaned = [
        " " if ord(char) < 32 or char in _INVALID_PORTABLE_FILENAME_CHARS else char
        for char in text
    ]
    text = re.sub(r"\s+", " ", "".join(cleaned)).strip(" .")
    return (text or fallback)[:120].rstrip(" .") or fallback


def sanitize_image_prefix(prefix, fallback=DEFAULT_IMAGE_PREFIX):
    """Return a Nalbantov-style one-to-four-character image prefix."""
    cleaned = "".join(
        char
        for char in str(prefix or "").upper()
        if "A" <= char <= "Z" or "0" <= char <= "9"
    )[:4]
    if cleaned:
        return cleaned
    fallback_cleaned = "".join(
        char
        for char in str(fallback or DEFAULT_IMAGE_PREFIX).upper()
        if "A" <= char <= "Z" or "0" <= char <= "9"
    )[:4]
    return fallback_cleaned or DEFAULT_IMAGE_PREFIX


def _numbered_image_stem(image_prefix, image_number):
    image_number = int(image_number)
    if not 0 <= image_number <= MAX_IMAGE_NUMBER:
        raise FloppyImageError(
            f"Disk number must be between 0 and {MAX_IMAGE_NUMBER}."
        )
    return f"{sanitize_image_prefix(image_prefix)}{image_number:04d}"


def _volume_label(image_prefix, image_number):
    return _numbered_image_stem(image_prefix, image_number)[:11]


def _raise_if_cancelled(cancel_callback):
    if cancel_callback is not None and cancel_callback():
        raise FloppyOperationCancelled("Emulator image creation cancelled.")


def _notify(progress_callback, step, total, message):
    if progress_callback is not None:
        progress_callback(int(step or 0), max(1, int(total or 1)), str(message or ""))


def _prepared_song_name(source_path, index, output_content):
    return build_dos83_filename(
        os.path.basename(source_path),
        index,
        extension="FIL" if output_content == "eseq" else "MID",
    )


def _prepare_song_files(
    song_paths,
    temp_directory,
    *,
    output_content,
    progress_callback=None,
    cancel_callback=None,
    midi_to_eseq_converter=None,
    eseq_to_midi_converter=None,
    language_code=None,
):
    midi_to_eseq_converter = midi_to_eseq_converter or convert_midi_file_to_eseq_path
    eseq_to_midi_converter = eseq_to_midi_converter or convert_eseq_file_to_midi_path
    prepared = []
    converted_count = 0
    total = len(song_paths)

    for index, source_path in enumerate(song_paths, start=1):
        _raise_if_cancelled(cancel_callback)
        source_name = os.path.basename(source_path)
        source_is_midi = is_midi_file(source_path)
        source_is_eseq = is_eseq_file(source_path)
        if not source_is_midi and not source_is_eseq:
            raise FloppyImageError(
                f"'{source_name}' is not a valid Standard MIDI or Yamaha E-SEQ file."
            )

        output_label = "E-SEQ" if output_content == "eseq" else "MIDI"
        _notify(
            progress_callback,
            index - 1,
            max(1, total * 2),
            tr(
                "emulator.progress.preparing_file",
                language_code,
                filename=source_name,
                format=output_label,
                current=index,
                total=total,
            ),
        )

        image_path = _prepared_song_name(source_path, index, output_content)
        local_path = os.path.join(
            temp_directory,
            f"{index:05d}_{uuid.uuid4().hex}_{image_path}",
        )
        try:
            if output_content == "eseq" and source_is_midi:
                midi_to_eseq_converter(
                    source_path,
                    local_path,
                    filename_hint=image_path,
                )
                converted_count += 1
            elif output_content == "eseq" and is_clavinova_mda_file(source_path):
                intermediate_midi = os.path.join(
                    temp_directory,
                    f"{index:05d}_{uuid.uuid4().hex}.mid",
                )
                eseq_to_midi_converter(source_path, intermediate_midi)
                midi_to_eseq_converter(
                    intermediate_midi,
                    local_path,
                    filename_hint=image_path,
                )
                converted_count += 1
            elif output_content == "eseq":
                update_error = update_eseq_order_key_to_path(
                    source_path,
                    build_eseq_order_key_from_path(image_path),
                    local_path,
                )
                if update_error:
                    raise FloppyImageError(update_error)
            elif source_is_eseq:
                eseq_to_midi_converter(source_path, local_path)
                converted_count += 1
            else:
                shutil.copy2(source_path, local_path)
        except Exception as exc:
            raise FloppyImageError(
                f"Could not prepare '{source_name}' as {output_label}: {exc}"
            ) from exc
        output_is_valid = (
            is_eseq_file(local_path)
            if output_content == "eseq"
            else is_midi_file(local_path)
        )
        if not os.path.isfile(local_path) or not output_is_valid:
            raise FloppyImageError(
                f"Preparation did not produce a valid {output_label} file for '{source_name}'."
            )

        # Describe what was actually written, not the source metadata. Native
        # MIDI titles remain full-length, while an E-SEQ output title reflects
        # the format's physical 32-byte field.
        title = (
            extract_first_title_from_midi(local_path)
            if output_content == "midi"
            else extract_eseq_title_from_file(local_path)
        )
        if title.startswith("Error"):
            title = ""
        prepared.append(
            _PreparedSong(
                source_path=source_path,
                image_path=image_path,
                local_path=local_path,
                title=title or os.path.splitext(source_name)[0],
            )
        )

    return prepared, converted_count


def _directory_path_for_songs(temp_directory, songs, metadata, image_number):
    track_entries = [
        PianodirTrackEntry(
            image_path=song.image_path,
            local_path=song.local_path,
            title=song.title,
        )
        for song in songs
    ]
    output_path = os.path.join(
        temp_directory,
        f"pianodir_{image_number:03d}_{uuid.uuid4().hex}.fil",
    )
    with open(output_path, "wb") as handle:
        handle.write(build_pianodir_bytes(track_entries, metadata=metadata))
    return output_path


def _metadata_for_disk(metadata, image_prefix, disk_number):
    """Use the emulator slot as the catalog ID unless an API caller overrides it."""
    catalog_number = (
        metadata.catalog_number
        or f"{sanitize_image_prefix(image_prefix)}-{int(disk_number):04d}"
    )
    return PianodirMetadata(
        catalog_number=catalog_number,
        disk_title=metadata.disk_title or catalog_number,
        raw_label_bytes=metadata.raw_label_bytes,
    )


def _verify_raw_image(
    raw_path,
    songs,
    *,
    include_pianodir,
    safety_margin_bytes=0,
):
    listing = read_image_listing(raw_path)
    actual_names = {
        os.path.basename(entry.path).upper()
        for entry in listing.entries
        if not entry.directory
    }
    expected_names = {song.image_path.upper() for song in songs}
    if include_pianodir:
        expected_names.add(PIANODIR_FILENAME)
    if actual_names != expected_names:
        missing = sorted(expected_names - actual_names)
        extra = sorted(actual_names - expected_names)
        detail = []
        if missing:
            detail.append(f"missing: {', '.join(missing)}")
        if extra:
            detail.append(f"unexpected: {', '.join(extra)}")
        raise FloppyImageError(
            "Generated image verification failed"
            + (f" ({'; '.join(detail)})" if detail else ".")
        )
    if listing.free_space < safety_margin_bytes:
        raise FloppyImageError(
            "Generated image verification failed: the disk has only "
            f"{listing.free_space:,} bytes free, below the requested "
            f"{safety_margin_bytes:,}-byte safety margin."
        )


def _pack_raw_images(
    prepared_songs,
    temp_directory,
    disk_format,
    *,
    image_prefix,
    starting_number,
    safety_margin_bytes,
    metadata,
    output_content,
    progress_callback=None,
    cancel_callback=None,
    language_code=None,
):
    include_pianodir = output_content == "eseq"
    placeholder_path = (
        _directory_path_for_songs(
            temp_directory,
            [],
            _metadata_for_disk(metadata, image_prefix, starting_number),
            0,
        )
        if include_pianodir
        else ""
    )
    raw_images = []
    current_raw = ""
    current_songs = []
    total = len(prepared_songs)

    def start_image(image_number):
        raw_path = os.path.join(temp_directory, f"raw_{image_number:03d}.img")
        disk_number = starting_number + image_number - 1
        if disk_number > MAX_IMAGE_NUMBER:
            raise FloppyImageError(
                f"This set would exceed disk {MAX_IMAGE_NUMBER}; choose a lower "
                "starting number."
            )
        create_blank_floppy_image(
            raw_path,
            disk_format,
            volume_label=_volume_label(image_prefix, disk_number),
            cancel_callback=cancel_callback,
        )
        if include_pianodir:
            _copy_host_file_into_image(
                raw_path,
                placeholder_path,
                PIANODIR_FILENAME,
                cancel_callback=cancel_callback,
            )
        return raw_path

    def has_room_with_margin(raw_path, song):
        listing = read_image_listing(raw_path)
        required_bytes = allocated_size(
            os.path.getsize(song.local_path),
            listing.cluster_size,
        )
        return listing.free_space - required_bytes >= safety_margin_bytes

    def too_large_message(song):
        preparation_note = (
            " after E-SEQ conversion" if output_content == "eseq" else ""
        )
        margin_note = (
            f" while preserving the {safety_margin_bytes // 1024:,} KiB safety margin"
            if safety_margin_bytes
            else ""
        )
        return (
            f"'{os.path.basename(song.source_path)}' is too large to fit on "
            f"a {disk_format.label} image{preparation_note}{margin_note}."
        )

    def finish_image():
        if not current_raw or not current_songs:
            return
        image_number = len(raw_images) + 1
        if include_pianodir:
            disk_number = starting_number + image_number - 1
            directory_path = _directory_path_for_songs(
                temp_directory,
                current_songs,
                _metadata_for_disk(metadata, image_prefix, disk_number),
                image_number,
            )
            _delete_eseq_directory_entries_from_image(
                current_raw,
                cancel_callback=cancel_callback,
            )
            _copy_host_file_into_image(
                current_raw,
                directory_path,
                PIANODIR_FILENAME,
                cancel_callback=cancel_callback,
            )
        _verify_raw_image(
            current_raw,
            current_songs,
            include_pianodir=include_pianodir,
            safety_margin_bytes=safety_margin_bytes,
        )
        raw_images.append((current_raw, tuple(current_songs)))

    current_raw = start_image(1)
    for index, song in enumerate(prepared_songs, start=1):
        _raise_if_cancelled(cancel_callback)
        _notify(
            progress_callback,
            total + index - 1,
            max(1, total * 2),
            tr(
                "emulator.progress.packing",
                language_code,
                filename=song.image_path,
                disk=len(raw_images) + 1,
            ),
        )

        if include_pianodir and len(current_songs) >= PIANODIR_MAX_TRACKS:
            finish_image()
            current_songs = []
            current_raw = start_image(len(raw_images) + 1)

        if not has_room_with_margin(current_raw, song):
            if not current_songs:
                raise FloppyImageError(too_large_message(song))
            finish_image()
            current_songs = []
            current_raw = start_image(len(raw_images) + 1)
            if not has_room_with_margin(current_raw, song):
                raise FloppyImageError(too_large_message(song))

        try:
            _copy_host_file_into_image(
                current_raw,
                song.local_path,
                song.image_path,
                cancel_callback=cancel_callback,
            )
            current_songs.append(song)
            continue
        except FloppyImageError as exc:
            if not _is_image_capacity_error(exc):
                raise

        if not current_songs:
            raise FloppyImageError(too_large_message(song))

        finish_image()
        current_songs = []
        current_raw = start_image(len(raw_images) + 1)
        if not has_room_with_margin(current_raw, song):
            raise FloppyImageError(too_large_message(song))
        try:
            _copy_host_file_into_image(
                current_raw,
                song.local_path,
                song.image_path,
                cancel_callback=cancel_callback,
            )
        except FloppyImageError as exc:
            if _is_image_capacity_error(exc):
                raise FloppyImageError(too_large_message(song)) from exc
            raise
        current_songs.append(song)

    finish_image()
    return raw_images


def _output_paths(
    output_directory,
    image_prefix,
    output_ext,
    image_count,
    starting_number,
):
    filenames = [
        f"{_numbered_image_stem(image_prefix, image_number)}.{output_ext}"
        for image_number in range(
            starting_number,
            starting_number + image_count,
        )
    ]
    return [os.path.join(output_directory, filename) for filename in filenames]


def _song_lists_output_path(
    output_directory,
    image_prefix,
    starting_number,
    image_count,
):
    first_stem = _numbered_image_stem(image_prefix, starting_number)
    if image_count == 1:
        filename = f"{first_stem}-song-list.txt"
    else:
        last_stem = _numbered_image_stem(
            image_prefix,
            starting_number + image_count - 1,
        )
        filename = f"{first_stem}-{last_stem}-song-lists.txt"
    return os.path.join(output_directory, filename)


def _song_list_display_text(value, fallback=""):
    text = re.sub(r"\s+", " ", str(value or "").replace("\x00", " ")).strip()
    return text or fallback


def _build_emulator_song_lists_text(
    raw_images,
    final_paths,
    *,
    image_prefix,
    starting_number,
    metadata,
    output_content,
):
    lines = [
        "Emulator Disk Set Song Lists",
        f"Images: {len(final_paths)}",
        f"Songs: {sum(len(songs) for _raw_path, songs in raw_images)}",
        "",
    ]
    for image_index, ((_raw_path, songs), final_path) in enumerate(
        zip(raw_images, final_paths),
        start=1,
    ):
        lines.append(
            f"Image {image_index} of {len(final_paths)}: {os.path.basename(final_path)}"
        )
        if output_content == "eseq":
            disk_metadata = _metadata_for_disk(
                metadata,
                image_prefix,
                starting_number + image_index - 1,
            )
            disk_title = _song_list_display_text(disk_metadata.disk_title)
            catalog_number = _song_list_display_text(disk_metadata.catalog_number)
            if disk_title:
                lines.append(f"Album: {disk_title}")
            if catalog_number:
                lines.append(f"Catalog: {catalog_number}")
        lines.append("")
        for song_index, song in enumerate(songs, start=1):
            lines.append(
                f"{song_index}. {_song_list_display_text(song.title, 'Untitled')}"
            )
        if image_index < len(final_paths):
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def build_emulator_disk_images(
    source_directory,
    output_directory,
    *,
    prefix=None,
    starting_number=DEFAULT_STARTING_NUMBER,
    safety_margin_bytes=DEFAULT_SAFETY_MARGIN_BYTES,
    set_name=None,
    album_title="",
    catalog_number="",
    disk_format=None,
    output_ext="hfe",
    output_content="eseq",
    include_subfolders=True,
    shuffle=False,
    include_song_lists=False,
    overwrite_existing=False,
    overwrite_callback=None,
    language_code=None,
    progress_callback=None,
    cancel_callback=None,
    midi_to_eseq_converter=None,
    eseq_to_midi_converter=None,
):
    """Build numbered emulator-ready images containing only E-SEQ or only MIDI.

    ``set_name`` is retained as a compatibility alias for ``prefix``.
    """
    source_directory = os.path.abspath(os.fspath(source_directory))
    output_directory = os.path.abspath(os.fspath(output_directory))
    output_ext = str(output_ext or "hfe").lower().lstrip(".")
    if output_ext not in EMULATOR_IMAGE_EXTENSIONS:
        raise FloppyImageError(
            "Emulator image format must be raw IMG or HFE."
        )
    output_content = str(output_content or "eseq").strip().lower()
    if output_content not in EMULATOR_CONTENT_FORMATS:
        raise FloppyImageError("Disk contents must be Yamaha E-SEQ or Standard MIDI.")
    if disk_format is None:
        disk_format = next(item for item in DISK_FORMATS if item.key == "ibm.720")
    if disk_format not in DISK_FORMATS:
        raise FloppyImageError("Choose a supported IBM floppy disk format.")
    try:
        starting_number = int(starting_number)
    except (TypeError, ValueError) as exc:
        raise FloppyImageError("The starting disk number must be a whole number.") from exc
    if not 0 <= starting_number <= MAX_IMAGE_NUMBER:
        raise FloppyImageError(
            f"The starting disk number must be between 0 and {MAX_IMAGE_NUMBER}."
        )
    try:
        safety_margin_bytes = int(safety_margin_bytes)
    except (TypeError, ValueError) as exc:
        raise FloppyImageError("The safety margin must be a whole number of bytes.") from exc
    if safety_margin_bytes < 0:
        raise FloppyImageError("The safety margin cannot be negative.")

    song_paths = discover_song_files(
        source_directory,
        include_subfolders=include_subfolders,
    )
    if not song_paths:
        raise FloppyImageError(
            "The selected folder does not contain any MIDI or Yamaha E-SEQ song files."
        )
    shuffle = bool(shuffle)
    include_song_lists = bool(include_song_lists)
    if shuffle:
        random.shuffle(song_paths)

    image_prefix = sanitize_image_prefix(prefix if prefix is not None else set_name)
    album_title = str(album_title or "").strip()
    catalog_number = str(catalog_number or "").strip()
    metadata = PianodirMetadata(
        catalog_number=catalog_number,
        disk_title=album_title,
    )
    try:
        os.makedirs(output_directory, exist_ok=True)
    except OSError as exc:
        raise FloppyImageError(f"Could not create the output folder: {exc}") from exc
    if not os.path.isdir(output_directory):
        raise FloppyImageError(f"The output path is not a folder: {output_directory}")

    temp_directory = tempfile.mkdtemp(prefix="aps_emulator_images_")
    committed_paths = []
    replacement_backups = {}
    try:
        prepared_songs, converted_count = _prepare_song_files(
            song_paths,
            temp_directory,
            output_content=output_content,
            progress_callback=progress_callback,
            cancel_callback=cancel_callback,
            midi_to_eseq_converter=midi_to_eseq_converter,
            eseq_to_midi_converter=eseq_to_midi_converter,
            language_code=language_code,
        )
        raw_images = _pack_raw_images(
            prepared_songs,
            temp_directory,
            disk_format,
            image_prefix=image_prefix,
            starting_number=starting_number,
            safety_margin_bytes=safety_margin_bytes,
            metadata=metadata,
            output_content=output_content,
            progress_callback=progress_callback,
            cancel_callback=cancel_callback,
            language_code=language_code,
        )
        ending_number = starting_number + len(raw_images) - 1
        if ending_number > MAX_IMAGE_NUMBER:
            raise FloppyImageError(
                f"This set would end at disk {ending_number}; choose a starting "
                f"number that keeps every disk at {MAX_IMAGE_NUMBER} or below."
            )
        final_paths = _output_paths(
            output_directory,
            image_prefix,
            output_ext,
            len(raw_images),
            starting_number,
        )
        song_list_path = (
            _song_lists_output_path(
                output_directory,
                image_prefix,
                starting_number,
                len(raw_images),
            )
            if include_song_lists
            else ""
        )
        output_candidates = [*final_paths]
        if song_list_path:
            output_candidates.append(song_list_path)
        existing_paths = [path for path in output_candidates if os.path.lexists(path)]
        if existing_paths:
            invalid_paths = [
                path
                for path in existing_paths
                if not os.path.isfile(path) or os.path.islink(path)
            ]
            if invalid_paths:
                names = ", ".join(
                    os.path.basename(path) for path in invalid_paths[:3]
                )
                if len(invalid_paths) > 3:
                    names += ", ..."
                raise FloppyImageError(
                    f"Output path cannot be replaced as a regular file: {names}."
                )

            overwrite_approved = bool(overwrite_existing)
            if not overwrite_approved and overwrite_callback is not None:
                _raise_if_cancelled(cancel_callback)
                overwrite_approved = bool(overwrite_callback(tuple(existing_paths)))
                _raise_if_cancelled(cancel_callback)
            if not overwrite_approved:
                if overwrite_callback is not None:
                    raise FloppyOperationCancelled(
                        "Emulator image replacement was cancelled."
                    )
                names = ", ".join(os.path.basename(path) for path in existing_paths[:3])
                if len(existing_paths) > 3:
                    names += ", ..."
                raise FloppyImageError(
                    f"Output file already exists: {names}. Choose another prefix, "
                    "starting number, or output folder."
                )

        staged_outputs = []
        total_steps = len(prepared_songs) * 2 + len(raw_images)
        for index, ((raw_path, _songs), final_path) in enumerate(
            zip(raw_images, final_paths),
            start=1,
        ):
            _raise_if_cancelled(cancel_callback)
            _notify(
                progress_callback,
                len(prepared_songs) * 2 + index - 1,
                total_steps,
                tr(
                    "emulator.progress.writing",
                    language_code,
                    current=index,
                    total=len(raw_images),
                ),
            )
            staged_path = os.path.join(
                temp_directory,
                f"output_{index:03d}.{output_ext}",
            )
            _write_image_direct(raw_path, staged_path, output_ext, disk_format)
            staged_outputs.append((staged_path, final_path))

        if song_list_path:
            staged_song_list_path = os.path.join(
                temp_directory,
                "emulator_song_lists.txt",
            )
            with open(
                staged_song_list_path,
                "w",
                encoding="utf-8",
                newline="\n",
            ) as handle:
                handle.write(
                    _build_emulator_song_lists_text(
                        raw_images,
                        final_paths,
                        image_prefix=image_prefix,
                        starting_number=starting_number,
                        metadata=metadata,
                        output_content=output_content,
                    )
                )
            staged_outputs.append((staged_song_list_path, song_list_path))

        _raise_if_cancelled(cancel_callback)
        for backup_index, existing_path in enumerate(existing_paths, start=1):
            backup_path = os.path.join(
                temp_directory,
                f"existing_output_{backup_index:04d}.bak",
            )
            shutil.copy2(existing_path, backup_path)
            replacement_backups[existing_path] = backup_path

        for staged_path, final_path in staged_outputs:
            committed_paths.append(final_path)
            _finish_temp_output(staged_path, final_path)

        _notify(
            progress_callback,
            total_steps,
            total_steps,
            tr(
                "emulator.progress.complete",
                language_code,
                images=len(final_paths),
            ),
        )
        return EmulatorImageBuildResult(
            source_directory=source_directory,
            output_directory=output_directory,
            song_files_found=len(song_paths),
            files_prepared=len(prepared_songs),
            converted_files=converted_count,
            images_created=len(final_paths),
            output_content=output_content,
            output_paths=tuple(final_paths),
            image_prefix=image_prefix,
            starting_number=starting_number,
            safety_margin_bytes=safety_margin_bytes,
            shuffled=shuffle,
            song_list_path=song_list_path,
        )
    except Exception:
        for path in reversed(committed_paths):
            try:
                backup_path = replacement_backups.get(path)
                if backup_path and os.path.isfile(backup_path):
                    shutil.copy2(backup_path, path)
                elif os.path.lexists(path):
                    os.remove(path)
            except OSError:
                pass
        raise
    finally:
        shutil.rmtree(temp_directory, ignore_errors=True)


def build_emulator_eseq_images(*args, converter=None, **kwargs):
    """Compatibility wrapper for the original E-SEQ-only builder API."""
    kwargs["output_content"] = "eseq"
    if converter is not None:
        kwargs["midi_to_eseq_converter"] = converter
    return build_emulator_disk_images(*args, **kwargs)
