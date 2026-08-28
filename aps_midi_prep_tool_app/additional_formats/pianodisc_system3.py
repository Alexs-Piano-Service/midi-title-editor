"""Read PianoDisc System 3 floppy images and convert their songs to MIDI.

System 3 disks do not contain a FAT filesystem.  Their directory is a table of
44-byte records and each song is stored as a compact, MIDI-like event stream in
one or more contiguous 512-byte sectors.  Timing values in that stream map
directly to a 480 PPQ Standard MIDI file at 120 BPM.

The implementation is deliberately read-only: it recovers playable Standard
MIDI files without attempting to author or modify PianoDisc media.
"""

from __future__ import annotations

import re
import struct
from dataclasses import dataclass


SECTOR_SIZE = 512
SYSTEM_SIGNATURE = b"1 PianoDisc"
SYSTEM_VERSION = b"System 3"
SYSTEM_VERSION_OFFSET = 0x20
CATALOG_RECORD_SIZE = 44
CATALOG_TITLE_SIZE = 22
CATALOG_ACTIVE_VALUE = 1
MIDI_DIVISION = 480
MIDI_TEMPO_US_PER_QUARTER = 500_000
END_PAUSE_TICKS = MIDI_DIVISION * 3


class PianoDiscSystem3Error(ValueError):
    """Raised when a System 3 image or song stream cannot be decoded safely."""


@dataclass(frozen=True)
class PianoDiscSong:
    number: int
    title: str
    start_sector: int
    sector_count: int
    data: bytes


@dataclass(frozen=True)
class PianoDiscMidiFile:
    song: PianoDiscSong
    filename: str
    data: bytes


@dataclass(frozen=True)
class PianoDiscImageConversion:
    files: tuple[PianoDiscMidiFile, ...]
    errors: tuple[str, ...]


@dataclass(frozen=True)
class _SourceEvent:
    delay: int
    status: int
    data: bytes
    sysex: bool = False


def _system_header_offset(data: bytes) -> int | None:
    search_limit = min(len(data), 64 * SECTOR_SIZE)
    offset = data.find(SYSTEM_SIGNATURE, 0, search_limit)
    while offset >= 0:
        # Real production disks do not consistently include the optional
        # "System 3" version text used by early fixtures. The sector-aligned
        # product signature plus the validated catalog is the stable format
        # identifier.
        if offset % SECTOR_SIZE == 0:
            return offset
        offset = data.find(SYSTEM_SIGNATURE, offset + 1, search_limit)
    return None


def looks_like_pianodisc_system3_bytes(data: bytes) -> bool:
    """Return whether *data* has a valid-looking PianoDisc System 3 catalog."""
    try:
        return bool(parse_pianodisc_system3_image(data))
    except PianoDiscSystem3Error:
        return False


def _decode_catalog_title(raw_title: bytes, number: int) -> str:
    raw_title = (
        raw_title.split(b"\x00", 1)[0]
        .replace(b"\xfe", b" ")
        .replace(b"\xff", b" ")
        .rstrip(b" ")
    )
    title = raw_title.decode("cp1252", errors="replace")
    title = "".join(" " if ord(char) < 32 else char for char in title)
    title = re.sub(r"\s+", " ", title).strip()
    return title or f"PianoDisc Track {number:02d}"


def _parse_pianodisc_system3_catalog(
    data: bytes,
) -> tuple[tuple[PianoDiscSong, ...], tuple[str, ...]]:
    """Parse valid catalog songs and retain per-record damage reports.

    System 3 catalog entries hold a sector count and starting sector.  Song
    extents are copied exactly as described by the catalog. A malformed active
    record is skipped rather than reading outside the image or hiding valid
    songs that follow it.
    """
    data = bytes(data)
    header_offset = _system_header_offset(data)
    if header_offset is None:
        raise PianoDiscSystem3Error("PianoDisc System 3 header not found.")
    if len(data) % SECTOR_SIZE:
        raise PianoDiscSystem3Error("The PianoDisc image is not sector-aligned.")

    catalog_offset = header_offset + SECTOR_SIZE
    if catalog_offset + CATALOG_RECORD_SIZE > len(data):
        raise PianoDiscSystem3Error("The PianoDisc catalog is truncated.")

    songs = []
    errors = []
    total_sectors = len(data) // SECTOR_SIZE
    max_records = min(256, (len(data) - catalog_offset) // CATALOG_RECORD_SIZE)
    for record_index in range(max_records):
        offset = catalog_offset + record_index * CATALOG_RECORD_SIZE
        record = data[offset:offset + CATALOG_RECORD_SIZE]
        active = int.from_bytes(record[22:24], "little")
        if active != CATALOG_ACTIVE_VALUE:
            break

        number = record_index + 1
        title = _decode_catalog_title(record[:CATALOG_TITLE_SIZE], number)
        sector_count = int.from_bytes(record[24:26], "little")
        start_sector = int.from_bytes(record[26:28], "little")
        if sector_count <= 0:
            errors.append(
                f"Track {number:02d} ({title}): catalog entry has no allocated sectors."
            )
            continue
        if start_sector <= 0 or start_sector + sector_count > total_sectors:
            errors.append(
                f"Track {number:02d} ({title}): catalog entry points outside the image."
            )
            continue

        start = start_sector * SECTOR_SIZE
        end = (start_sector + sector_count) * SECTOR_SIZE
        songs.append(
            PianoDiscSong(
                number=number,
                title=title,
                start_sector=start_sector,
                sector_count=sector_count,
                data=data[start:end],
            )
        )

    if not songs:
        detail = errors[0] if errors else "The PianoDisc System 3 catalog is empty or invalid."
        raise PianoDiscSystem3Error(detail)
    return tuple(songs), tuple(errors)


def parse_pianodisc_system3_image(data: bytes) -> tuple[PianoDiscSong, ...]:
    """Parse the catalog and return valid songs in catalog order."""
    songs, _errors = _parse_pianodisc_system3_catalog(data)
    return songs


def _read_source_delay(stream: bytes, offset: int) -> tuple[int | None, int]:
    if offset >= len(stream):
        raise PianoDiscSystem3Error("Song data ended while reading a timing value.")
    first = stream[offset]
    if first == 0xFD:
        if offset + 4 > len(stream):
            raise PianoDiscSystem3Error("Song data has a truncated 24-bit timing value.")
        return int.from_bytes(stream[offset + 1:offset + 4], "little"), offset + 4
    if first == 0xFE:
        if offset + 3 > len(stream):
            raise PianoDiscSystem3Error("Song data has a truncated 16-bit timing value.")
        return int.from_bytes(stream[offset + 1:offset + 3], "little"), offset + 3
    if first == 0xFF:
        # System 3 uses FF as a no-op/padding token between event records.
        return None, offset + 1
    return first, offset + 1


def _channel_data_length(status: int) -> int:
    kind = status & 0xF0
    if kind in (0xC0, 0xD0):
        return 1
    if kind in (0x80, 0x90, 0xA0, 0xB0, 0xE0):
        return 2
    raise PianoDiscSystem3Error(f"Unsupported PianoDisc event status 0x{status:02X}.")


def _iter_source_events(stream: bytes):
    offset = 0
    running_status = None
    found_end = False

    while offset < len(stream):
        delay, offset = _read_source_delay(stream, offset)
        if delay is None:
            continue
        if offset >= len(stream):
            raise PianoDiscSystem3Error("Song data ended before its next event.")

        first = stream[offset]
        if first == 0xFC:
            found_end = True
            yield _SourceEvent(delay=delay, status=first, data=b"")
            break

        if first == 0xF0:
            end = stream.find(b"\xF7", offset + 1)
            if end < 0:
                raise PianoDiscSystem3Error("PianoDisc SysEx data has no F7 terminator.")
            yield _SourceEvent(
                delay=delay,
                status=0xF0,
                data=stream[offset + 1:end + 1],
                sysex=True,
            )
            offset = end + 1
            running_status = None
            continue

        if first & 0x80:
            running_status = first
            offset += 1
            first_data = None
        else:
            if running_status is None:
                raise PianoDiscSystem3Error(
                    f"Running-status data 0x{first:02X} has no preceding MIDI status."
                )
            first_data = first
            offset += 1

        data_length = _channel_data_length(running_status)
        event_data = bytearray()
        if first_data is not None:
            event_data.append(first_data)
        needed = data_length - len(event_data)
        if offset + needed > len(stream):
            raise PianoDiscSystem3Error("Song data contains a truncated MIDI event.")
        event_data.extend(stream[offset:offset + needed])
        offset += needed
        if any(value & 0x80 for value in event_data):
            raise PianoDiscSystem3Error("Song data contains an invalid MIDI data byte.")
        yield _SourceEvent(
            delay=delay,
            status=running_status,
            data=bytes(event_data),
        )

    if not found_end:
        raise PianoDiscSystem3Error("Song data has no System 3 end marker.")


def _encode_vlq(value: int) -> bytes:
    value = max(0, int(value))
    encoded = bytearray([value & 0x7F])
    value >>= 7
    while value:
        encoded.insert(0, 0x80 | (value & 0x7F))
        value >>= 7
    return bytes(encoded)


def _meta_event(delta: int, meta_type: int, payload: bytes) -> bytes:
    return _encode_vlq(delta) + bytes((0xFF, meta_type)) + _encode_vlq(len(payload)) + payload


def _channel_event(delta: int, status: int, payload: bytes) -> bytes:
    return _encode_vlq(delta) + bytes((status,)) + payload


def _safe_midi_filename(number: int, title: str) -> str:
    cleaned = "".join(" " if char in '<>:"/\\|?*' or ord(char) < 32 else char for char in title)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    if not cleaned:
        cleaned = f"PianoDisc Track {number:02d}"
    return f"{number:02d} - {cleaned[:120].rstrip(' .')}.mid"


def _dos83_midi_filename(number: int) -> str:
    """Return the conventional short PianoDisc export name for a track."""
    if number < 1:
        raise ValueError("Track number must be at least 1.")
    if number <= 999:
        return f"PIANO{number:03d}.MID"
    if number <= 9_999_999:
        return f"P{number:07d}.MID"
    raise ValueError("Track number is too large for a DOS 8.3 filename.")


def convert_pianodisc_song_to_midi(song: PianoDiscSong) -> bytes:
    """Convert one decoded System 3 song to an SMF0 piano performance."""
    track = bytearray()
    track.extend(
        _meta_event(
            0,
            0x51,
            MIDI_TEMPO_US_PER_QUARTER.to_bytes(3, "big"),
        )
    )
    track.extend(_meta_event(0, 0x58, bytes((4, 2, 24, 8))))
    track.extend(_meta_event(0, 0x59, bytes((0, 0))))
    track.extend(_meta_event(0, 0x03, song.title.encode("utf-8", errors="replace")))
    track.extend(
        _meta_event(
            0,
            0x02,
            b"Converted from PianoDisc System 3 floppy",
        )
    )
    track.extend(_channel_event(0, 0xC0, b"\x00"))

    pending_note_on = None
    marked_note = 0
    active_notes = [False] * 128
    accumulated_delay = 0
    reached_end = False

    def write_event(delta: int, event: _SourceEvent):
        if event.sysex:
            track.extend(_encode_vlq(delta))
            track.append(0xF0)
            track.extend(_encode_vlq(len(event.data)))
            track.extend(event.data)
        else:
            track.extend(_channel_event(delta, event.status, event.data))

    for event in _iter_source_events(song.data):
        is_cancel_pending = (
            event.status & 0xF0 == 0xB0
            and len(event.data) == 2
            and event.data[0] == 0x16
        )
        if pending_note_on is not None and not is_cancel_pending:
            write_event(accumulated_delay, pending_note_on)
            active_notes[pending_note_on.data[0]] = True
            accumulated_delay = 0
            pending_note_on = None

        accumulated_delay += event.delay
        if event.status == 0xFC:
            reached_end = True
            break

        kind = event.status & 0xF0
        if kind == 0xB0 and len(event.data) == 2:
            controller, value = event.data
            if controller == 0x15:
                marked_note = value
                continue
            if controller == 0x16:
                pending_note_on = None
                continue

        if kind == 0x90 and len(event.data) == 2 and event.data[1] > 0:
            pending_note_on = event
            continue

        if kind in (0x80, 0x90) and len(event.data) == 2:
            note = event.data[0]
            if note == marked_note:
                marked_note = 0
                continue
            if not active_notes[note]:
                continue
            active_notes[note] = False

        write_event(accumulated_delay, event)
        accumulated_delay = 0

    if not reached_end:
        raise PianoDiscSystem3Error("Song conversion ended without a System 3 end marker.")

    track.extend(_meta_event(accumulated_delay + END_PAUSE_TICKS, 0x2F, b""))
    header = b"MThd" + struct.pack(">IHHH", 6, 0, 1, MIDI_DIVISION)
    return header + b"MTrk" + struct.pack(">I", len(track)) + bytes(track)


def convert_pianodisc_system3_image(
    data: bytes,
    *,
    long_filenames: bool = False,
) -> PianoDiscImageConversion:
    """Convert every valid catalog song and report damaged entries separately."""
    songs, catalog_errors = _parse_pianodisc_system3_catalog(data)
    files = []
    errors = list(catalog_errors)
    used_names = set()
    for song in songs:
        try:
            midi_data = convert_pianodisc_song_to_midi(song)
        except PianoDiscSystem3Error as exc:
            errors.append(f"Track {song.number:02d} ({song.title}): {exc}")
            continue

        filename = (
            _safe_midi_filename(song.number, song.title)
            if long_filenames
            else _dos83_midi_filename(song.number)
        )
        stem, extension = filename.rsplit(".", 1)
        candidate = filename
        suffix = 2
        while candidate.casefold() in used_names:
            candidate = f"{stem} ({suffix}).{extension}"
            suffix += 1
        used_names.add(candidate.casefold())
        files.append(PianoDiscMidiFile(song=song, filename=candidate, data=midi_data))

    if not files:
        detail = errors[0] if errors else "No catalog songs could be decoded."
        raise PianoDiscSystem3Error(detail)
    return PianoDiscImageConversion(files=tuple(files), errors=tuple(errors))
