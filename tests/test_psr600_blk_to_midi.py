import tempfile
import unittest
from pathlib import Path

from aps_midi_prep_tool_app.additional_formats import psr600_blk_to_midi
from aps_midi_prep_tool_app.midi_metadata import (
    extract_first_title_from_midi,
)
from aps_midi_prep_tool_app.midi_type0_converter import (
    _parse_midi_chunks,
    _parse_track_events,
)


def _sample_page_memory(apparent_layer_banks=()):
    apparent_layer_banks = {
        int(number) for number in apparent_layer_banks
    }
    data = bytearray(psr600_blk_to_midi.PAGE_MEMORY_SIZE)
    data[:len(psr600_blk_to_midi.PAGE_MEMORY_SIGNATURE)] = (
        psr600_blk_to_midi.PAGE_MEMORY_SIGNATURE
    )
    data[psr600_blk_to_midi.TEMPO_OFFSET] = 120
    data[psr600_blk_to_midi.STYLE_OFFSET] = 7

    for index in range(5):
        bank_start = (
            psr600_blk_to_midi.MELODY_START
            + index * psr600_blk_to_midi.MELODY_SIZE
        )
        data[bank_start + 4] = 0xF2

    bank_start = psr600_blk_to_midi.MELODY_START
    data[bank_start + 4] = 0xF1
    setup = bytes([
        0,
        0,
        0,
        64,
        35,
        2,
        111,
        48,
        31,
        0xFF,
        0,
        2,
        106,
        72,
        71,
        0xFF,
    ])
    bank_one_setup = bytearray(setup)
    if 1 in apparent_layer_banks:
        bank_one_setup[0] = 0x7F
    data[bank_start + 5:bank_start + 21] = bank_one_setup
    data[bank_start + 21:bank_start + 29] = bytes([
        0x94,
        60,
        100,
        96,
        0x84,
        60,
        0xF2,
        0,
    ])

    bank_start += psr600_blk_to_midi.MELODY_SIZE
    data[bank_start + 4] = 0xF1
    bank_two_setup = bytearray(setup)
    if 2 in apparent_layer_banks:
        bank_two_setup[0] = 0x7F
    data[bank_start + 5:bank_start + 21] = bank_two_setup
    data[bank_start + 21:bank_start + 29] = bytes([
        0x95,
        64,
        90,
        48,
        0x85,
        64,
        0xF2,
        0,
    ])
    return bytes(data)


class Psr600BlkToMidiTests(unittest.TestCase):
    def test_recognizes_and_summarizes_page_memory(self):
        data = _sample_page_memory()
        self.assertTrue(
            psr600_blk_to_midi.looks_like_psr600_blk_bytes(data)
        )
        self.assertFalse(
            psr600_blk_to_midi.looks_like_psr600_blk_bytes(data[:-1])
        )

        page = psr600_blk_to_midi.parse_page_memory_bytes(
            data,
            "SAMPLE.BLK",
        )
        self.assertEqual(page.tempo_bpm, 120)
        self.assertEqual(page.style_number, 7)
        self.assertEqual(len(page.used_melody_banks), 2)
        self.assertEqual(page.used_melody_banks[0].initial_voice, 35)
        self.assertEqual(page.used_melody_banks[0].secondary_voice, 0)
        self.assertFalse(
            page.used_melody_banks[0].apparent_layer_enabled
        )
        self.assertEqual(page.used_melody_banks[0].note_on_count, 1)
        self.assertEqual(page.used_melody_banks[1].note_on_count, 1)
        self.assertIsNone(page.used_melody_banks[0].error)
        self.assertIsNone(page.used_melody_banks[1].error)

    def test_groups_recorded_banks_in_one_type_one_midi(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "SAMPLE.BLK"
            source.write_bytes(_sample_page_memory())
            reports = psr600_blk_to_midi.convert_one(
                source,
                Path(temp_dir) / "midi",
            )

            self.assertEqual(len(reports), 1)
            self.assertEqual(reports[0].melody_bank_count, 2)
            self.assertEqual(reports[0].note_on_count, 2)
            self.assertFalse(reports[0].partial)
            self.assertEqual(Path(reports[0].output).name, "SAMPLE.mid")
            self.assertEqual(
                extract_first_title_from_midi(reports[0].output),
                "SAMPLE",
            )

            midi_bytes = Path(reports[0].output).read_bytes()
            _header_end, format_type, track_count, chunks = (
                _parse_midi_chunks(midi_bytes)
            )
            track_chunks = [
                chunk for chunk in chunks if chunk["id"] == b"MTrk"
            ]
            self.assertEqual(format_type, 1)
            self.assertEqual(track_count, 3)
            self.assertEqual(len(track_chunks), 3)

            track_note_events = []
            track_end_ticks = []
            for track_chunk in track_chunks:
                events, last_tick = _parse_track_events(
                    midi_bytes[
                        track_chunk["data_start"]:track_chunk["data_end"]
                    ]
                )
                track_note_events.append([
                    (tick, raw)
                    for tick, _order, raw in events
                    if raw and (raw[0] & 0xF0) in (0x80, 0x90)
                ])
                track_end_ticks.append(last_tick)
            self.assertEqual(
                track_note_events,
                [
                    [],
                    [
                        (0, bytes([0x94, 60, 100])),
                        (96, bytes([0x84, 60, 0])),
                    ],
                    [
                        (0, bytes([0x95, 64, 90])),
                        (48, bytes([0x85, 64, 0])),
                    ],
                ],
            )
            self.assertEqual(track_end_ticks, [0, 96, 48])

    def test_exports_apparent_dual_voice_as_questioned_layer_track(self):
        page = psr600_blk_to_midi.parse_page_memory_bytes(
            _sample_page_memory(apparent_layer_banks={1}),
            "LAYER.BLK",
        )
        self.assertEqual(len(page.apparent_layer_banks), 1)
        layer_bank = page.apparent_layer_banks[0]
        self.assertEqual(layer_bank.apparent_layer_flag, 0x7F)
        self.assertEqual(
            layer_bank.primary_voice_descriptor,
            bytes([35, 2, 111, 48, 31, 0xFF]),
        )
        self.assertEqual(
            layer_bank.secondary_voice_descriptor,
            bytes([0, 2, 106, 72, 71, 0xFF]),
        )

        midi_bytes = psr600_blk_to_midi.build_multitrack_midi_bytes(
            page
        )
        _header_end, format_type, track_count, chunks = (
            _parse_midi_chunks(midi_bytes)
        )
        track_chunks = [
            chunk for chunk in chunks if chunk["id"] == b"MTrk"
        ]
        self.assertEqual(format_type, 1)
        self.assertEqual(track_count, 4)
        self.assertEqual(len(track_chunks), 4)

        parsed_tracks = []
        for track_chunk in track_chunks:
            events, _last_tick = _parse_track_events(
                midi_bytes[
                    track_chunk["data_start"]:track_chunk["data_end"]
                ]
            )
            parsed_tracks.append(events)

        primary_events = parsed_tracks[1]
        layer_events = parsed_tracks[2]
        next_primary_events = parsed_tracks[3]
        primary_title = next(
            raw[3:].decode("utf-8")
            for _tick, _order, raw in primary_events
            if raw[:2] == b"\xFF\x03"
        )
        layer_title = next(
            raw[3:].decode("utf-8")
            for _tick, _order, raw in layer_events
            if raw[:2] == b"\xFF\x03"
        )
        self.assertIn("PSR 35 Strings 1 + PSR 00 Piano?", primary_title)
        self.assertIn("Layer? - PSR 00 Piano", layer_title)
        self.assertTrue(any(
            b"unconfirmed setup flag[0]=0x7F" in raw
            for _tick, _order, raw in layer_events
        ))

        primary_messages = [
            raw
            for _tick, _order, raw in primary_events
            if raw and 0x80 <= raw[0] <= 0xEF
        ]
        layer_messages = [
            raw
            for _tick, _order, raw in layer_events
            if raw and 0x80 <= raw[0] <= 0xEF
        ]
        next_primary_messages = [
            raw
            for _tick, _order, raw in next_primary_events
            if raw and 0x80 <= raw[0] <= 0xEF
        ]
        self.assertIn(
            bytes([0xC4, psr600_blk_to_midi.gm_program(35)]),
            primary_messages,
        )
        self.assertIn(bytes([0xCA, 0]), layer_messages)
        self.assertIn(bytes([0x9A, 60, 100]), layer_messages)
        self.assertIn(bytes([0x8A, 60, 0]), layer_messages)
        self.assertIn(
            bytes([0xC5, psr600_blk_to_midi.gm_program(35)]),
            next_primary_messages,
        )
        self.assertFalse(any(
            raw[0] & 0x0F == 4
            for raw in layer_messages
        ))


if __name__ == "__main__":
    unittest.main()
