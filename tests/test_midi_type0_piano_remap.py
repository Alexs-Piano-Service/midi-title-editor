import struct
import unittest

from aps_midi_prep_tool_app.midi_type0_converter import (
    _convert_midi_bytes_to_type0,
    _encode_vlq,
    _parse_midi_chunks,
    _parse_track_events,
)


def _track(events, end_tick=0):
    payload = bytearray()
    previous_tick = 0
    for tick, raw in events:
        payload.extend(_encode_vlq(tick - previous_tick))
        payload.extend(raw)
        previous_tick = tick
    payload.extend(_encode_vlq(end_tick - previous_tick))
    payload.extend(b"\xFF\x2F\x00")
    return b"MTrk" + len(payload).to_bytes(4, "big") + bytes(payload)


def _type1_midi(*tracks):
    return struct.pack(">4sIHHH", b"MThd", 6, 1, len(tracks), 480) + b"".join(tracks)


class MidiType0PianoRemapTests(unittest.TestCase):
    def _events(self, midi_bytes):
        _header_end, format_type, track_count, chunks = _parse_midi_chunks(midi_bytes)
        self.assertEqual(format_type, 0)
        self.assertEqual(track_count, 1)
        track = next(chunk for chunk in chunks if chunk["id"] == b"MTrk")
        events, _end_tick = _parse_track_events(
            midi_bytes[track["data_start"]:track["data_end"]]
        )
        return events

    def test_remaps_all_channel_events_and_selects_acoustic_grand(self):
        source = _type1_midi(
            _track(
                [
                    (0, bytes([0xC2, 40])),
                    (0, bytes([0x92, 60, 100])),
                    (10, bytes([0x82, 60, 0])),
                    (10, bytes([0xB2, 64, 90])),
                    (20, bytes([0xB2, 123, 0])),
                ],
                end_tick=20,
            ),
            _track(
                [
                    (0, bytes([0xB9, 0, 1])),
                    (0, bytes([0xC9, 12])),
                    (5, bytes([0x99, 36, 110])),
                    (15, bytes([0x89, 36, 0])),
                ],
                end_tick=20,
            ),
        )

        converted, changed = _convert_midi_bytes_to_type0(
            source,
            remap_all_instruments_to_channel0=True,
        )

        self.assertTrue(changed)
        events = self._events(converted)
        channel_events = [raw for _tick, _order, raw in events if 0x80 <= raw[0] <= 0xEF]
        self.assertTrue(channel_events)
        self.assertTrue(all((raw[0] & 0x0F) == 0 for raw in channel_events))
        self.assertEqual([raw for raw in channel_events if (raw[0] & 0xF0) == 0xC0], [b"\xC0\x00"])
        self.assertIn(bytes([0x90, 60, 100]), channel_events)
        self.assertIn(bytes([0x90, 36, 110]), channel_events)
        self.assertIn(bytes([0xB0, 64, 90]), channel_events)
        self.assertNotIn(bytes([0xB0, 0, 1]), channel_events)
        self.assertNotIn(bytes([0xB0, 123, 0]), channel_events)
        note_events = [
            (tick, raw)
            for tick, _order, raw in events
            if (raw[0] & 0xF0) in (0x80, 0x90)
        ]
        self.assertEqual(
            note_events,
            [
                (0, bytes([0x90, 60, 100])),
                (5, bytes([0x90, 36, 110])),
                (10, bytes([0x80, 60, 0])),
                (15, bytes([0x80, 36, 0])),
            ],
        )

    def test_default_type0_conversion_preserves_source_channels(self):
        source = _type1_midi(
            _track([(0, bytes([0x92, 60, 100])), (10, bytes([0x82, 60, 0]))], end_tick=10),
            _track([(0, bytes([0x95, 67, 100])), (10, bytes([0x85, 67, 0]))], end_tick=10),
        )

        converted, changed = _convert_midi_bytes_to_type0(source)

        self.assertTrue(changed)
        channels = {
            raw[0] & 0x0F
            for _tick, _order, raw in self._events(converted)
            if raw and 0x80 <= raw[0] <= 0xEF
        }
        self.assertEqual(channels, {2, 5})


if __name__ == "__main__":
    unittest.main()
