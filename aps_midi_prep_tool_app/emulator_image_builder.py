"""Build emulator-ready MIDI or Yamaha E-SEQ floppy-image sets."""

from __future__ import annotations

import os
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


def _volume_label(set_name, image_number):
    base = "".join(
        char for char in str(set_name or "").upper()
        if "A" <= char <= "Z" or "0" <= char <= "9"
    ) or "ESEQ"
    suffix = f"{max(1, int(image_number)):02d}"
    return f"{base[:max(1, 11 - len(suffix))]}{suffix}"[:11]


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

        title = (
            extract_first_title_from_midi(source_path)
            if source_is_midi
            else extract_eseq_title_from_file(source_path)
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


def _verify_raw_image(raw_path, songs, *, include_pianodir):
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


def _pack_raw_images(
    prepared_songs,
    temp_directory,
    disk_format,
    *,
    set_name,
    metadata,
    output_content,
    progress_callback=None,
    cancel_callback=None,
    language_code=None,
):
    include_pianodir = output_content == "eseq"
    placeholder_path = (
        _directory_path_for_songs(temp_directory, [], metadata, 0)
        if include_pianodir
        else ""
    )
    raw_images = []
    current_raw = ""
    current_songs = []
    total = len(prepared_songs)

    def start_image(image_number):
        raw_path = os.path.join(temp_directory, f"raw_{image_number:03d}.img")
        create_blank_floppy_image(
            raw_path,
            disk_format,
            volume_label=_volume_label(set_name, image_number),
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

    def finish_image():
        if not current_raw or not current_songs:
            return
        image_number = len(raw_images) + 1
        if include_pianodir:
            directory_path = _directory_path_for_songs(
                temp_directory,
                current_songs,
                metadata,
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
            preparation_note = (
                " after E-SEQ conversion" if output_content == "eseq" else ""
            )
            raise FloppyImageError(
                f"'{os.path.basename(song.source_path)}' is too large to fit on "
                f"a {disk_format.label} image{preparation_note}."
            )

        finish_image()
        current_songs = []
        current_raw = start_image(len(raw_images) + 1)
        try:
            _copy_host_file_into_image(
                current_raw,
                song.local_path,
                song.image_path,
                cancel_callback=cancel_callback,
            )
        except FloppyImageError as exc:
            if _is_image_capacity_error(exc):
                preparation_note = (
                    " after E-SEQ conversion" if output_content == "eseq" else ""
                )
                raise FloppyImageError(
                    f"'{os.path.basename(song.source_path)}' is too large to fit on "
                    f"a {disk_format.label} image{preparation_note}."
                ) from exc
            raise
        current_songs.append(song)

    finish_image()
    return raw_images


def _output_paths(output_directory, set_name, output_ext, image_count):
    if image_count == 1:
        filenames = [f"{set_name}.{output_ext}"]
    else:
        digits = max(2, len(str(image_count)))
        filenames = [
            f"{set_name}_{index:0{digits}d}.{output_ext}"
            for index in range(1, image_count + 1)
        ]
    return [os.path.join(output_directory, filename) for filename in filenames]


def build_emulator_disk_images(
    source_directory,
    output_directory,
    *,
    set_name="",
    album_title="",
    catalog_number="",
    disk_format=None,
    output_ext="img",
    output_content="eseq",
    include_subfolders=True,
    language_code=None,
    progress_callback=None,
    cancel_callback=None,
    midi_to_eseq_converter=None,
    eseq_to_midi_converter=None,
):
    """Build emulator-ready disk images containing only E-SEQ or only MIDI."""
    source_directory = os.path.abspath(os.fspath(source_directory))
    output_directory = os.path.abspath(os.fspath(output_directory))
    output_ext = str(output_ext or "img").lower().lstrip(".")
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

    song_paths = discover_song_files(
        source_directory,
        include_subfolders=include_subfolders,
    )
    if not song_paths:
        raise FloppyImageError(
            "The selected folder does not contain any MIDI or Yamaha E-SEQ song files."
        )

    set_name = sanitize_image_set_name(
        set_name or os.path.basename(os.path.normpath(source_directory))
    )
    album_title = str(album_title or set_name).strip()
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
            set_name=set_name,
            metadata=metadata,
            output_content=output_content,
            progress_callback=progress_callback,
            cancel_callback=cancel_callback,
            language_code=language_code,
        )
        final_paths = _output_paths(
            output_directory,
            set_name,
            output_ext,
            len(raw_images),
        )
        existing_paths = [path for path in final_paths if os.path.lexists(path)]
        if existing_paths:
            names = ", ".join(os.path.basename(path) for path in existing_paths[:3])
            if len(existing_paths) > 3:
                names += ", ..."
            raise FloppyImageError(
                f"Output image already exists: {names}. Choose another set name "
                "or output folder."
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

        _raise_if_cancelled(cancel_callback)
        for staged_path, final_path in staged_outputs:
            _finish_temp_output(staged_path, final_path)
            committed_paths.append(final_path)

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
        )
    except Exception:
        for path in committed_paths:
            try:
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
