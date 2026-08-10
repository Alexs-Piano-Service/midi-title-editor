import os
import re
import unicodedata


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
_MAX_FILENAME_LENGTH = 240


def _clean_filename_title(title):
    text = unicodedata.normalize("NFC", str(title or "")).replace("\x00", " ")
    text = _INVALID_FILENAME_CHARS.sub("-", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text


def _truncate_utf8(text, max_bytes):
    encoded = text.encode("utf-8")
    if len(encoded) <= max_bytes:
        return text
    encoded = encoded[:max_bytes]
    while encoded:
        try:
            return encoded.decode("utf-8")
        except UnicodeDecodeError:
            encoded = encoded[:-1]
    return ""


def build_long_midi_filename(track_number, title, source_filename="", track_count=None):
    """Build a portable, descriptive MIDI filename from track order and title."""
    track_number = int(track_number)
    if track_number < 1:
        raise ValueError("Track number must be at least 1.")

    count = max(track_number, int(track_count or 0))
    number_width = max(2, len(str(count)))
    prefix = f"{track_number:0{number_width}d} - "

    clean_title = _clean_filename_title(title)
    if not clean_title:
        source_stem = os.path.splitext(os.path.basename(source_filename or ""))[0]
        clean_title = _clean_filename_title(source_stem) or "Untitled"

    max_title_bytes = max(1, _MAX_FILENAME_LENGTH - len(prefix.encode("utf-8")) - len(".mid"))
    clean_title = _truncate_utf8(clean_title, max_title_bytes).rstrip(" .") or "Untitled"
    return f"{prefix}{clean_title}.mid"
