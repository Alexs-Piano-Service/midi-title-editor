"""Read Yamaha Smart PianoSoft song titles from PSONG.MNG catalogs."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass

from .midi_metadata import normalize_title_spacing


SMART_PIANOSOFT_SONG_CATALOG_NAME = "PSONG.MNG"
_CATALOG_HEADER_SIZE = 0x80
_CATALOG_RECORD_SIZE = 0xB0


@dataclass(frozen=True)
class SmartPianoSoftSong:
    track_number: int
    filename: str
    title: str


def parse_smart_pianosoft_song_catalog(data):
    """Parse PSONG.MNG and return its songs in catalog order."""
    data = bytes(data)
    if len(data) < _CATALOG_HEADER_SIZE:
        raise ValueError("PSONG.MNG is truncated.")
    if not data[:14].upper().startswith(b"PSONG   MNG"):
        raise ValueError("PSONG.MNG header not found.")

    count_match = re.match(rb"FILE(\d{3})", data[0x20:0x30].upper())
    if count_match is None:
        raise ValueError("PSONG.MNG song count not found.")
    record_count = int(count_match.group(1))
    required_size = _CATALOG_HEADER_SIZE + record_count * _CATALOG_RECORD_SIZE
    if required_size > len(data):
        raise ValueError("PSONG.MNG song records are truncated.")

    songs = []
    for record_index in range(record_count):
        start = _CATALOG_HEADER_SIZE + record_index * _CATALOG_RECORD_SIZE
        record = data[start:start + _CATALOG_RECORD_SIZE]
        stem = record[:8].decode("cp1252", errors="replace").strip(" \x00")
        extension = record[8:11].decode("cp1252", errors="replace").strip(" \x00")
        if not stem:
            continue
        filename = f"{stem}.{extension}" if extension else stem
        raw_title = record[0x10:0x30].decode("cp1252", errors="replace")
        songs.append(
            SmartPianoSoftSong(
                track_number=record_index + 1,
                filename=filename,
                title=normalize_title_spacing(raw_title),
            )
        )
    return tuple(songs)


def smart_pianosoft_catalog_by_filename(data):
    """Return a case-insensitive filename lookup for a PSONG.MNG payload."""
    return {
        song.filename.casefold(): song
        for song in parse_smart_pianosoft_song_catalog(data)
    }


def update_smart_pianosoft_song_catalog(data, title_updates):
    """Return PSONG.MNG bytes with titles replaced by source filename."""
    songs = parse_smart_pianosoft_song_catalog(data)
    updates = {
        os.path.basename(str(filename or "")).casefold(): str(title or "")
        for filename, title in dict(title_updates or {}).items()
    }
    known_filenames = {song.filename.casefold() for song in songs}
    unknown = sorted(set(updates) - known_filenames)
    if unknown:
        raise ValueError(f"PSONG.MNG has no song record for {unknown[0]}.")

    patched = bytearray(data)
    for song in songs:
        new_title = updates.get(song.filename.casefold())
        if new_title is None:
            continue
        try:
            encoded_title = new_title.encode("cp1252")
        except UnicodeEncodeError as exc:
            raise ValueError("Smart PianoSoft titles must use Windows-1252 characters.") from exc
        if len(encoded_title) > 32:
            raise ValueError("Smart PianoSoft titles must be 32 bytes or fewer.")
        record_start = _CATALOG_HEADER_SIZE + (song.track_number - 1) * _CATALOG_RECORD_SIZE
        patched[record_start + 0x10:record_start + 0x30] = encoded_title.ljust(32, b" ")
    return bytes(patched)


def update_smart_pianosoft_catalog_to_path(source_path, destination_path, title_updates):
    """Write a PSONG.MNG copy with the requested catalog title updates."""
    with open(source_path, "rb") as handle:
        payload = update_smart_pianosoft_song_catalog(handle.read(), title_updates)
    with open(destination_path, "wb") as handle:
        handle.write(payload)


def smart_pianosoft_catalog_from_session(session, entries):
    """Read PSONG.MNG from an image/floppy session, or return an empty lookup."""
    for entry in entries:
        if (
            os.path.basename(str(entry.path or "")).casefold()
            != SMART_PIANOSOFT_SONG_CATALOG_NAME.casefold()
        ):
            continue
        try:
            catalog_path = session.extract_file(entry.path)
            with open(catalog_path, "rb") as handle:
                return smart_pianosoft_catalog_by_filename(handle.read())
        except Exception:
            return {}
    return {}
