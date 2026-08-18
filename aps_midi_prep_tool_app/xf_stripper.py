import os
import uuid


_SYSTEM_MESSAGE_DATA_LENGTHS = {
    0xF1: 1,
    0xF2: 2,
    0xF3: 1,
    0xF6: 0,
    0xF8: 0,
    0xF9: 0,
    0xFA: 0,
    0xFB: 0,
    0xFC: 0,
    0xFD: 0,
    0xFE: 0,
}


def _read_vlq(data, offset):
    value = 0
    for _ in range(4):
        if offset >= len(data):
            raise ValueError("Invalid MIDI variable-length value.")
        byte = data[offset]
        offset += 1
        value = (value << 7) | (byte & 0x7F)
        if not byte & 0x80:
            return value, offset
    raise ValueError("Invalid MIDI variable-length value.")


def _encode_vlq(value):
    if value < 0 or value > 0x0FFFFFFF:
        raise ValueError("MIDI variable-length value is out of range.")
    encoded = [value & 0x7F]
    value >>= 7
    while value:
        encoded.append(0x80 | (value & 0x7F))
        value >>= 7
    return bytes(reversed(encoded))


def _take(data, offset, length):
    end = offset + length
    if end > len(data):
        raise ValueError("A MIDI event is truncated.")
    return data[offset:end], end


def _strip_xf_track(track_data):
    offset = 0
    running_status = None
    pending_delta = 0
    removed_events = 0
    output = bytearray()

    while offset < len(track_data):
        delta, offset = _read_vlq(track_data, offset)
        pending_delta += delta
        if offset >= len(track_data):
            raise ValueError("A MIDI event is truncated.")

        first = track_data[offset]
        status_from_stream = bool(first & 0x80)
        if status_from_stream:
            status = first
            offset += 1
            if status < 0xF0:
                running_status = status
            elif status < 0xF8:
                running_status = None
        else:
            if running_status is None:
                raise ValueError("Invalid MIDI running status.")
            status = running_status

        if status == 0xFF:
            if not status_from_stream or offset >= len(track_data):
                raise ValueError("A MIDI meta event is truncated.")
            meta_type = track_data[offset]
            offset += 1
            meta_length, offset = _read_vlq(track_data, offset)
            payload, offset = _take(track_data, offset, meta_length)
            if meta_type == 0x7F:
                removed_events += 1
                continue
            if meta_type == 0x2F:
                break
            output.extend(_encode_vlq(pending_delta))
            output.extend((0xFF, meta_type))
            output.extend(_encode_vlq(meta_length))
            output.extend(payload)
            pending_delta = 0
            continue

        if status in (0xF0, 0xF7):
            if not status_from_stream:
                raise ValueError("Invalid MIDI running status.")
            sysex_length, offset = _read_vlq(track_data, offset)
            payload, offset = _take(track_data, offset, sysex_length)
            output.extend(_encode_vlq(pending_delta))
            output.append(status)
            output.extend(_encode_vlq(sysex_length))
            output.extend(payload)
            pending_delta = 0
            continue

        if 0x80 <= status <= 0xEF:
            message_type = status & 0xF0
            data_length = 1 if message_type in (0xC0, 0xD0) else 2
            payload, offset = _take(track_data, offset, data_length)
            output.extend(_encode_vlq(pending_delta))
            output.append(status)
            output.extend(payload)
            pending_delta = 0
            continue

        if not status_from_stream:
            raise ValueError("Invalid MIDI running status.")
        data_length = _SYSTEM_MESSAGE_DATA_LENGTHS.get(status)
        if data_length is None:
            raise ValueError(f"Unsupported MIDI status 0x{status:02X}.")
        payload, offset = _take(track_data, offset, data_length)
        output.extend(_encode_vlq(pending_delta))
        output.append(status)
        output.extend(payload)
        pending_delta = 0

    output.extend(_encode_vlq(pending_delta))
    output.extend(b"\xFF\x2F\x00")
    return bytes(output), removed_events


def strip_xf_from_midi_bytes(midi_bytes):
    """Remove Yamaha XF metadata while preserving Standard MIDI performance data."""
    if len(midi_bytes) < 14 or midi_bytes[:4] != b"MThd":
        raise ValueError("This is not a valid Standard MIDI File.")

    header_length = int.from_bytes(midi_bytes[4:8], "big")
    header_end = 8 + header_length
    if header_length < 6 or header_end > len(midi_bytes):
        raise ValueError("The MIDI header is invalid or truncated.")

    format_type = int.from_bytes(midi_bytes[8:10], "big")
    track_count = int.from_bytes(midi_bytes[10:12], "big")
    division = int.from_bytes(midi_bytes[12:14], "big")
    if format_type not in (0, 1, 2) or track_count < 1:
        raise ValueError("The MIDI header contains unsupported values.")
    if division & 0x8000:
        raise ValueError("SMPTE-timed MIDI files are not supported.")

    output = bytearray(b"MThd")
    output.extend((6).to_bytes(4, "big"))
    output.extend(format_type.to_bytes(2, "big"))
    output.extend(track_count.to_bytes(2, "big"))
    output.extend(division.to_bytes(2, "big"))

    offset = header_end
    for _track_index in range(track_count):
        if offset + 8 > len(midi_bytes) or midi_bytes[offset:offset + 4] != b"MTrk":
            raise ValueError("A declared MIDI track is missing or malformed.")
        track_length = int.from_bytes(midi_bytes[offset + 4:offset + 8], "big")
        track_start = offset + 8
        track_end = track_start + track_length
        if track_end > len(midi_bytes):
            raise ValueError("A MIDI track is truncated.")
        track_data, _removed_events = _strip_xf_track(midi_bytes[track_start:track_end])
        output.extend(b"MTrk")
        output.extend(len(track_data).to_bytes(4, "big"))
        output.extend(track_data)
        offset = track_end

    stripped_bytes = bytes(output)
    return stripped_bytes, stripped_bytes != midi_bytes


def strip_xf_from_midi_path(source_path, dest_path):
    """Strip XF data into ``dest_path`` and return whether any bytes changed."""
    if not os.path.isfile(source_path):
        raise ValueError("File does not exist.")

    with open(source_path, "rb") as handle:
        midi_bytes = handle.read()

    stripped_bytes, changed = strip_xf_from_midi_bytes(midi_bytes)
    if not changed:
        return False

    temp_path = f"{dest_path}.aps_xf_{uuid.uuid4().hex}.tmp"
    try:
        with open(temp_path, "wb") as handle:
            handle.write(stripped_bytes)
        os.replace(temp_path, dest_path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

    return True
