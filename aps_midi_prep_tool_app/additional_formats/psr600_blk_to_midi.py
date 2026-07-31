"""Best-effort Yamaha PSR-600 Page Memory (``.BLK``) conversion.

PSR-600 BLK files are 49,152-byte Page Memory snapshots.  A snapshot can
contain five recorded Melody banks as well as Chord banks, Conductor data,
panel settings, Multi Pads, styles, and other proprietary state.

This module deliberately implements a second-tier preservation workflow:

* recognize PSR-600 Page Memory files conservatively;
* decode recorded Melody banks;
* write one dependency-free Type 1 MIDI per BLK, with one primary track per
  Melody bank and a cautious extra track for apparent Dual Voice layers;
* use approximate General MIDI programs so the result is easy to audition.

Chord/accompaniment generation, Conductor order, Multi Pads, custom styles,
unconfirmed layer semantics, and other proprietary state are not fully
rendered.  The original BLK file remains the preservation copy.

The layout and event decoding were derived from the PSR600_recovery reference
archive supplied for APS MIDI Prep Tool development.  They are reverse
engineered rather than based on a published Yamaha file specification.
"""

from __future__ import annotations

import re
import struct
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


PAGE_MEMORY_SIZE = 49_152
PAGE_MEMORY_SIGNATURE = b"\xF0\x43\x76\x0A\x01\x45"
MELODY_START = 0x0AD6
MELODY_SIZE = 0x13C9
MELODY_CONTINUATION_OFFSET = 0x0F05
CHORD_START = 0x6DC4
CHORD_SIZE = 0x0306
TEMPO_OFFSET = 0x0102
STYLE_OFFSET = 0x0103
BAD_SECTOR_MARKER = b"-=[BAD SECTOR]=-"
TICKS_PER_QUARTER = 96


# PSR-600 panel voices, indexed 0..99.
VOICE_NAMES = (
    "Piano", "Flange Piano", "Honky-Tonk Piano", "Electric Piano 1",
    "Electric Piano 2", "Electric Piano 3", "Harpsichord 1", "Harpsichord 2",
    "Glockenspiel", "Celesta", "Pipe Organ 1", "Pipe Organ 2",
    "Electronic Organ 1", "Electronic Organ 2", "Electronic Organ 3", "Electronic Organ 4",
    "Accordion 1", "Accordion 2", "Electric Guitar 1", "Electric Guitar 2",
    "Electric Guitar 3", "Tremolo Guitar", "Electric 12-String Guitar", "Distortion Guitar",
    "Jazz Guitar", "Jazz Guitar Octave", "Mute Guitar", "Mute Guitar Echo",
    "Steel Guitar", "Folk Guitar", "12-String Guitar", "Gut Guitar",
    "Violin 1", "Violin 2", "Cello", "Strings 1",
    "Strings 2", "Orchestra Hit", "Harp", "Banjo",
    "Vibraphone", "Marimba", "Steel Drum", "Trumpet",
    "Mute Trumpet 1", "Mute Trumpet 2", "Mute Trumpet 3", "Trombone",
    "Flugelhorn", "Horn", "Tuba", "Brass Ensemble 1",
    "Brass Ensemble 2", "Piccolo", "Flute", "Clarinet",
    "Bass Clarinet", "Oboe", "English Horn", "Bassoon",
    "Soprano Sax", "Alto Sax", "Tenor Sax", "Baritone Sax",
    "Ocarina", "Pan Flute", "Recorder", "Harmonica",
    "Samba Whistle", "Sax Ensemble 1", "Sax Ensemble 2", "Woodwind Ensemble",
    "Chorus", "Synth Lead", "Synth Brass", "Synth Strings",
    "Synth Tom", "Fantasy 1", "Fantasy 2", "Fantasy 3",
    "Bell Strings", "Seq Pad", "Electric Bass 1", "Electric Bass 2",
    "Fretless Bass", "Mute Bass", "Mute Bass Echo", "Slap Bass",
    "Wood Bass 1", "Wood Bass 2", "Synth Bass 1", "Synth Bass 2",
    "Synth Bass 3", "Bowed Bass", "Scratch w/Pitch", "Kick & Snare w/Pitch",
    "Tom w/Pitch", "Latin Percs w/Pitch", "Percussions w/Gate", "Percussions",
)


# Approximate zero-based General MIDI programs.  These are auditioning aids,
# not exact representations of the PSR-600 tone generator.
GM_PROGRAM_MAP = (
    0, 4, 3, 4, 5, 4, 6, 6, 9, 8, 19, 19, 16, 17, 18, 16,
    21, 23, 27, 26, 29, 27, 27, 30, 26, 26, 28, 28, 25, 25, 25, 24,
    40, 40, 42, 48, 49, 55, 46, 105, 11, 12, 114, 56, 59, 59, 59, 57,
    56, 60, 58, 61, 61, 72, 73, 71, 71, 68, 69, 70, 64, 65, 66, 67,
    79, 75, 74, 22, 78, 65, 66, 68, 52, 80, 62, 50, 118, 88, 89, 91,
    94, 95, 33, 34, 35, 33, 33, 36, 32, 32, 38, 39, 38, 43, 120, 118,
    117, 115, 118, 118,
)


@dataclass(frozen=True)
class ParsedEvent:
    tick: int
    status: int
    data: tuple[int, ...]
    source_offset: int


@dataclass
class MelodyBank:
    number: int
    used: bool
    header: bytes = b""
    setup: bytes = b""
    events: list[ParsedEvent] = field(default_factory=list)
    end_offset: Optional[int] = None
    error: Optional[str] = None

    @property
    def initial_voice(self) -> Optional[int]:
        return self.setup[4] if len(self.setup) >= 5 else None

    @property
    def primary_voice_descriptor(self) -> bytes:
        return self.setup[4:10] if len(self.setup) >= 10 else b""

    @property
    def secondary_voice_descriptor(self) -> bytes:
        return self.setup[10:16] if len(self.setup) >= 16 else b""

    @property
    def secondary_voice(self) -> Optional[int]:
        descriptor = self.secondary_voice_descriptor
        return descriptor[0] if descriptor else None

    @property
    def apparent_layer_flag(self) -> Optional[int]:
        return self.setup[0] if self.setup else None

    @property
    def apparent_layer_enabled(self) -> bool:
        return (
            len(self.setup) >= 16
            and self.apparent_layer_flag == 0x7F
            and self.secondary_voice is not None
        )

    @property
    def channel(self) -> int:
        counts = Counter(
            event.status & 0x0F
            for event in self.events
            if event.status < 0xF0
            and (event.status >> 4) in (0x8, 0x9, 0xB, 0xC, 0xD, 0xE)
        )
        return counts.most_common(1)[0][0] if counts else min(15, 3 + self.number)

    @property
    def layer_channel(self) -> int:
        # Melody banks normally occupy zero-based channels 4..8.  Keep
        # audition layers on 10..14, clear of GM percussion channel 9.
        return min(15, 9 + self.number)

    @property
    def note_on_count(self) -> int:
        return sum(
            1
            for event in self.events
            if event.status >> 4 == 0x9
            and len(event.data) == 2
            and event.data[1] > 0
        )

    @property
    def last_tick(self) -> int:
        return max((event.tick for event in self.events), default=0)


@dataclass
class PageMemory:
    source_name: str
    data: bytes
    tempo_bpm: int
    style_number: int
    bad_sectors: list[int]
    chord_banks_used: list[int]
    melody_banks: list[MelodyBank]

    @property
    def used_melody_banks(self) -> list[MelodyBank]:
        return [bank for bank in self.melody_banks if bank.used]

    @property
    def apparent_layer_banks(self) -> list[MelodyBank]:
        return [
            bank
            for bank in self.used_melody_banks
            if bank.apparent_layer_enabled
        ]


@dataclass
class ConversionReport:
    input: str
    output: str
    melody_bank_count: int
    note_on_count: int
    partial_bank_count: int
    apparent_layer_count: int = 0
    warnings: list[str] = field(default_factory=list)

    @property
    def partial(self) -> bool:
        return self.partial_bank_count > 0


def voice_name(program: Optional[int]) -> str:
    if program is None:
        return ""
    if 0 <= program < len(VOICE_NAMES):
        return VOICE_NAMES[program]
    return f"Undefined/Special {program}"


def gm_program(psr_program: int) -> int:
    if 0 <= psr_program < len(GM_PROGRAM_MAP):
        return GM_PROGRAM_MAP[psr_program]
    return max(0, min(127, psr_program))


def _logical_melody_block(block: bytes) -> bytes:
    continuation_end = MELODY_CONTINUATION_OFFSET + 4
    if (
        len(block) >= continuation_end
        and block[
            MELODY_CONTINUATION_OFFSET:
            MELODY_CONTINUATION_OFFSET + 2
        ] == b"\x04\xC1"
    ):
        return (
            block[:MELODY_CONTINUATION_OFFSET]
            + block[continuation_end:]
        )
    return block


def parse_melody_bank(raw_block: bytes, number: int) -> MelodyBank:
    block = _logical_melody_block(raw_block)
    bank = MelodyBank(number=number, used=False, header=block[:4])
    if len(block) < 5:
        bank.error = "truncated bank header"
        return bank

    position = 4
    marker = block[position]
    if marker == 0xF2:
        bank.end_offset = position
        return bank
    if marker != 0xF1:
        bank.used = True
        bank.error = (
            "expected F1/F2 at logical offset 0x4; "
            f"found 0x{marker:02X}"
        )
        return bank

    bank.used = True
    if len(block) < position + 17:
        bank.error = "truncated initial voice-parameter record"
        return bank
    bank.setup = block[position + 1:position + 17]
    position += 17
    tick = 0

    while position < len(block):
        if block[position] < 0x80:
            tick += block[position]
            position += 1
            if position >= len(block):
                bank.error = "delta time at end of bank"
                break

        status_offset = position
        status = block[position]
        position += 1
        if status < 0x80:
            bank.error = (
                f"expected status at 0x{status_offset:X}; "
                f"found 0x{status:02X}"
            )
            break

        high = status >> 4
        if high == 0x8:
            data_length = 1
        elif high == 0x9:
            data_length = 2
        elif high == 0xA:
            data_length = 1
        elif high == 0xB:
            data_length = 2
        elif high == 0xC:
            data_length = 1
        elif high == 0xD:
            data_length = 0
        elif high == 0xE:
            data_length = 1
        elif status == 0xF2:
            bank.events.append(ParsedEvent(tick, status, (), status_offset))
            bank.end_offset = status_offset
            break
        elif status == 0xF5:
            data_length = 0
        else:
            bank.error = (
                f"unknown status 0x{status:02X} at 0x{status_offset:X}"
            )
            break

        if position + data_length > len(block):
            bank.error = (
                f"truncated data for status 0x{status:02X} "
                f"at 0x{status_offset:X}"
            )
            break
        event_data = tuple(block[position:position + data_length])
        if any(value >= 0x80 for value in event_data):
            bank.error = (
                f"invalid high-bit data after 0x{status:02X} "
                f"at 0x{status_offset:X}"
            )
            break
        position += data_length
        bank.events.append(
            ParsedEvent(tick, status, event_data, status_offset)
        )

    return bank


def looks_like_psr600_blk_bytes(data: bytes) -> bool:
    """Return whether *data* has the stable PSR-600 Page Memory structure."""
    if (
        len(data) != PAGE_MEMORY_SIZE
        or not data.startswith(PAGE_MEMORY_SIGNATURE)
    ):
        return False
    bank_markers = [
        data[MELODY_START + index * MELODY_SIZE + 4]
        for index in range(5)
    ]
    return sum(marker in (0xF1, 0xF2) for marker in bank_markers) >= 3


def looks_like_psr600_blk_file(path: Path | str) -> bool:
    path = Path(path)
    try:
        if path.suffix.lower() != ".blk" or path.stat().st_size != PAGE_MEMORY_SIZE:
            return False
        return looks_like_psr600_blk_bytes(path.read_bytes())
    except OSError:
        return False


def parse_page_memory_bytes(data: bytes, source_name: str = "Page Memory") -> PageMemory:
    if len(data) != PAGE_MEMORY_SIZE:
        raise ValueError(
            f"{source_name}: expected {PAGE_MEMORY_SIZE} bytes, got {len(data)}"
        )
    if not data.startswith(PAGE_MEMORY_SIGNATURE):
        raise ValueError(
            f"{source_name}: missing Yamaha PSR-600 Page Memory signature"
        )

    tempo = data[TEMPO_OFFSET]
    if not 30 <= tempo <= 300:
        tempo = 120

    bad_sectors = [
        index
        for index in range(len(data) // 512)
        if BAD_SECTOR_MARKER in data[index * 512:(index + 1) * 512]
    ]

    chord_banks_used = []
    for number in range(1, 6):
        start = CHORD_START + (number - 1) * CHORD_SIZE
        block = data[start:start + CHORD_SIZE]
        if len(block) >= 5 and block[4] == 0xF1:
            chord_banks_used.append(number)

    melody_banks = []
    for number in range(1, 6):
        start = MELODY_START + (number - 1) * MELODY_SIZE
        melody_banks.append(
            parse_melody_bank(data[start:start + MELODY_SIZE], number)
        )

    return PageMemory(
        source_name=source_name,
        data=data,
        tempo_bpm=tempo,
        style_number=data[STYLE_OFFSET],
        bad_sectors=bad_sectors,
        chord_banks_used=chord_banks_used,
        melody_banks=melody_banks,
    )


def parse_page_memory_file(path: Path | str) -> PageMemory:
    path = Path(path)
    return parse_page_memory_bytes(path.read_bytes(), source_name=path.name)


def summarize_file(path: Path | str) -> dict:
    page = parse_page_memory_file(path)
    used = page.used_melody_banks
    return {
        "file_count": 1,
        "melody_bank_count": len(used),
        "complete_melody_bank_count": sum(not bank.error for bank in used),
        "partial_melody_bank_count": sum(bool(bank.error) for bank in used),
        "note_on_count": sum(bank.note_on_count for bank in used),
        "apparent_layer_count": len(page.apparent_layer_banks),
        "chord_bank_count": len(page.chord_banks_used),
        "bad_sector_count": len(page.bad_sectors),
        "filenames": [Path(path).name],
    }


def _encode_vlq(value: int) -> bytes:
    value = max(0, int(value))
    encoded = bytearray([value & 0x7F])
    value >>= 7
    while value:
        encoded.insert(0, 0x80 | (value & 0x7F))
        value >>= 7
    return bytes(encoded)


def _chunk(chunk_type: bytes, payload: bytes) -> bytes:
    return chunk_type + struct.pack(">I", len(payload)) + payload


def _meta_event(meta_type: int, payload: bytes) -> bytes:
    return b"\xFF" + bytes([meta_type]) + _encode_vlq(len(payload)) + payload


def _text_bytes(text: str) -> bytes:
    return text.encode("utf-8", "replace")


def _channel_event_bytes(
    event: ParsedEvent,
    use_gm: bool,
    *,
    channel_override: Optional[int] = None,
    include_program_changes: bool = True,
) -> Optional[bytes]:
    status = event.status
    data = event.data
    high = status >> 4
    channel = (
        status & 0x0F
        if channel_override is None
        else max(0, min(int(channel_override), 15))
    )

    if high == 0x8 and len(data) == 1:
        return bytes([0x80 | channel, data[0], 0])
    if high == 0x9 and len(data) == 2:
        return bytes([0x90 | channel, data[0], data[1]])
    if high == 0xA:
        return None
    if high == 0xB and len(data) == 2:
        return bytes([0xB0 | channel, data[0], data[1]])
    if high == 0xC and len(data) == 1:
        if not include_program_changes:
            return None
        program = gm_program(data[0]) if use_gm else data[0]
        return bytes([0xC0 | channel, program & 0x7F])
    if high == 0xD:
        return bytes([0xB0 | channel, 123, 0])
    if high == 0xE and len(data) == 1:
        pitch = max(-8192, min(8191, (data[0] - 64) * 128))
        encoded_pitch = pitch + 8192
        return bytes([
            0xE0 | channel,
            encoded_pitch & 0x7F,
            (encoded_pitch >> 7) & 0x7F,
        ])
    return None


def _voice_descriptor_messages(
    descriptor: bytes,
    channel: int,
    *,
    use_gm: bool,
) -> tuple[bytes, ...]:
    if len(descriptor) < 5:
        return ()
    program = descriptor[0]
    if use_gm:
        program = gm_program(program)
    return (
        bytes([0xC0 | channel, program & 0x7F]),
        bytes([0xB0 | channel, 7, descriptor[2] & 0x7F]),
        bytes([0xB0 | channel, 91, descriptor[3] & 0x7F]),
        bytes([0xB0 | channel, 10, descriptor[4] & 0x7F]),
        bytes([0xE0 | channel, 0, 64]),
    )


def _build_melody_track_bytes(
    page: PageMemory,
    bank: MelodyBank,
    *,
    use_gm: bool = True,
    layer: bool = False,
) -> bytes:
    """Build a primary or apparent-layer track for one Melody bank."""
    if not bank.used:
        raise ValueError(f"Melody Bank {bank.number} is empty")
    if layer and not bank.apparent_layer_enabled:
        raise ValueError(
            f"Melody Bank {bank.number} has no apparent Dual Voice layer"
        )

    voice = bank.secondary_voice if layer else bank.initial_voice
    descriptor = (
        bank.secondary_voice_descriptor
        if layer
        else bank.primary_voice_descriptor
    )
    channel = bank.layer_channel if layer else bank.channel
    title = (
        f"Melody Bank {bank.number} Layer?"
        if layer
        else f"Melody Bank {bank.number} Primary"
    )
    if voice is not None:
        title += f" - PSR {voice:02d} {voice_name(voice)}"
    if (
        not layer
        and bank.apparent_layer_enabled
        and bank.secondary_voice is not None
    ):
        title += (
            f" + PSR {bank.secondary_voice:02d} "
            f"{voice_name(bank.secondary_voice)}?"
        )
    if bank.error:
        title += " [PARTIAL]"

    track = bytearray()
    track.extend(b"\x00" + _meta_event(0x03, _text_bytes(title)))
    track.extend(
        b"\x00"
        + _meta_event(
            0x01,
            _text_bytes(
                f"Converted from Yamaha PSR-600 Page Memory: {page.source_name}; "
                f"Melody Bank {bank.number}; "
                + (
                    "apparent secondary voice layer"
                    if layer
                    else "primary voice"
                )
            ),
        )
    )
    track.extend(
        b"\x00"
        + _meta_event(
            0x01,
            _text_bytes(
                "Approximate General MIDI instruments; accompaniment, "
                "Conductor, Multi Pad, style, and unconfirmed proprietary "
                "settings remain only in the source BLK file."
                if use_gm
                else
                "Original PSR-600 program numbers; accompaniment, Conductor, "
                "Multi Pad, style, and unconfirmed proprietary settings "
                "remain only in the source BLK file."
            ),
        )
    )
    if bank.setup:
        layer_flag = bank.apparent_layer_flag
        layer_interpretation = (
            "treated as an apparent Dual Voice enable for auditioning"
            if bank.apparent_layer_enabled
            else "not treated as an enabled layer"
        )
        track.extend(
            b"\x00"
            + _meta_event(
                0x01,
                _text_bytes(
                    f"Primary voice descriptor: "
                    f"{bank.primary_voice_descriptor.hex(' ')}; "
                    f"secondary voice descriptor: "
                    f"{bank.secondary_voice_descriptor.hex(' ')}; "
                    f"unconfirmed setup flag[0]=0x{layer_flag:02X}, "
                    f"{layer_interpretation}; raw setup "
                    f"{bank.setup.hex(' ')}"
                ),
            )
        )
    if layer:
        track.extend(
            b"\x00"
            + _meta_event(
                0x01,
                _text_bytes(
                    "AUDITIONING INTERPRETATION: notes and controls duplicate "
                    f"Melody Bank {bank.number} on MIDI channel {channel + 1}; "
                    "recorded program changes are omitted so the stored "
                    "secondary voice remains selected. Dual Voice activation "
                    "has not been confirmed from a published BLK specification."
                ),
            )
        )
    if bank.error:
        track.extend(
            b"\x00"
            + _meta_event(
                0x01,
                _text_bytes(f"PARTIAL CONVERSION: {bank.error}"),
            )
        )

    scheduled: list[tuple[int, int, bytes]] = []
    sequence = 0
    for message in _voice_descriptor_messages(
        descriptor,
        channel,
        use_gm=use_gm,
    ):
        scheduled.append((0, sequence, message))
        sequence += 1

    for event in bank.events:
        if event.status in (0xF2, 0xF5):
            continue
        message = _channel_event_bytes(
            event,
            use_gm,
            channel_override=(channel if layer else None),
            include_program_changes=not layer,
        )
        if message is None:
            continue
        scheduled.append((event.tick, sequence, message))
        sequence += 1

    if bank.error:
        scheduled.append(
            (
                bank.last_tick + 1,
                sequence,
                bytes([0xB0 | channel, 123, 0]),
            )
        )

    scheduled.sort(key=lambda item: (item[0], item[1]))
    previous_tick = 0
    for tick, _sequence, message in scheduled:
        track.extend(_encode_vlq(max(0, tick - previous_tick)))
        track.extend(message)
        previous_tick = tick
    track.extend(b"\x00\xFF\x2F\x00")
    return bytes(track)


def _build_conductor_track_bytes(
    page: PageMemory,
    *,
    use_gm: bool = True,
) -> bytes:
    source_stem = Path(page.source_name).stem or "PSR-600 Page Memory"
    track = bytearray()
    track.extend(b"\x00" + _meta_event(0x03, _text_bytes(source_stem)))
    track.extend(
        b"\x00"
        + _meta_event(
            0x01,
            _text_bytes(
                f"Converted from Yamaha PSR-600 Page Memory: "
                f"{page.source_name}; {len(page.used_melody_banks)} recorded "
                "Melody banks are separate primary MIDI tracks; "
                f"{len(page.apparent_layer_banks)} apparent Dual Voice "
                "layer(s) are separate audition tracks."
            ),
        )
    )
    track.extend(
        b"\x00"
        + _meta_event(
            0x01,
            _text_bytes(
                "All Melody-bank tracks are aligned at tick zero. The "
                "PSR-600 Conductor sequence and Chord-bank Melody switching "
                "are not decoded, so sequential sections may overlap."
            ),
        )
    )
    track.extend(
        b"\x00"
        + _meta_event(
            0x01,
            _text_bytes(
                "Approximate General MIDI instruments; accompaniment, "
                "Conductor, Multi Pad, style, and unconfirmed proprietary "
                "settings remain only in the source BLK file."
                if use_gm
                else
                "Original PSR-600 program numbers; accompaniment, Conductor, "
                "Multi Pad, style, and unconfirmed proprietary settings "
                "remain only in the source BLK file."
            ),
        )
    )
    microseconds_per_quarter = max(
        1,
        min(0xFFFFFF, round(60_000_000 / page.tempo_bpm)),
    )
    track.extend(
        b"\x00"
        + _meta_event(
            0x51,
            microseconds_per_quarter.to_bytes(3, "big"),
        )
    )
    track.extend(b"\x00" + _meta_event(0x58, bytes([4, 2, 24, 8])))
    track.extend(b"\x00\xFF\x2F\x00")
    return bytes(track)


def build_multitrack_midi_bytes(
    page: PageMemory,
    *,
    use_gm: bool = True,
) -> bytes:
    """Build one Type 1 SMF containing every recorded Melody bank."""
    banks = page.used_melody_banks
    if not banks:
        raise ValueError("No recorded Melody banks were found.")

    track_specs = [
        (bank, layer)
        for bank in banks
        for layer in (
            (False, True)
            if bank.apparent_layer_enabled
            else (False,)
        )
    ]
    header = struct.pack(
        ">HHH",
        1,
        len(track_specs) + 1,
        TICKS_PER_QUARTER,
    )
    chunks = [
        _chunk(
            b"MTrk",
            _build_conductor_track_bytes(page, use_gm=use_gm),
        )
    ]
    chunks.extend(
        _chunk(
            b"MTrk",
            _build_melody_track_bytes(
                page,
                bank,
                use_gm=use_gm,
                layer=layer,
            ),
        )
        for bank, layer in track_specs
    )
    return _chunk(b"MThd", header) + b"".join(chunks)


def _safe_output_stem(value: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", value).strip(" ._")
    return stem or "PSR600_PAGE"


def convert_one(
    input_path: Path | str,
    output_dir: Path | str,
    *,
    output_stem: Optional[str] = None,
    use_gm: bool = True,
) -> list[ConversionReport]:
    """Convert one BLK file to one multitrack Type 1 MIDI."""
    input_path = Path(input_path)
    output_dir = Path(output_dir)
    page = parse_page_memory_file(input_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = _safe_output_stem(output_stem or input_path.stem)
    page.source_name = f"{stem}.BLK"

    banks = page.used_melody_banks
    if not banks:
        return []
    output_path = output_dir / f"{stem}.mid"
    output_path.write_bytes(
        build_multitrack_midi_bytes(page, use_gm=use_gm)
    )

    warnings = []
    if page.bad_sectors:
        sector_text = ", ".join(str(index) for index in page.bad_sectors)
        warnings.append(
            f"source contains damaged-sector marker(s) in sector(s) {sector_text}"
        )
    for bank in banks:
        if bank.error:
            warnings.append(f"Melody Bank {bank.number}: {bank.error}")
    return [
        ConversionReport(
            input=str(input_path),
            output=str(output_path),
            melody_bank_count=len(banks),
            note_on_count=sum(bank.note_on_count for bank in banks),
            partial_bank_count=sum(bool(bank.error) for bank in banks),
            apparent_layer_count=len(page.apparent_layer_banks),
            warnings=warnings,
        )
    ]
