import unittest

from aps_midi_prep_tool_app.main_window import (
    FluidSynthPlaybackProcess,
    MidiOutputWorker,
    _midi_meta_payload,
    _midi_channel_color,
    _midi_output_events,
    _midi_tick_for_seconds,
    _normalized_tempo_percent,
    _scale_midi_tempo_bytes,
    _scale_preview_timed_items,
)
from aps_midi_prep_tool_app.midi_type0_converter import (
    _encode_vlq,
    _parse_midi_chunks,
    _parse_track_events,
)


def _type_zero_midi(events):
    track = bytearray()
    previous_tick = 0
    for tick, raw in events:
        track.extend(_encode_vlq(tick - previous_tick))
        track.extend(raw)
        previous_tick = tick
    track.extend(b"\x00\xFF\x2F\x00")
    return (
        b"MThd"
        + (6).to_bytes(4, "big")
        + (0).to_bytes(2, "big")
        + (1).to_bytes(2, "big")
        + (96).to_bytes(2, "big")
        + b"MTrk"
        + len(track).to_bytes(4, "big")
        + bytes(track)
    )


def _track_events(midi_bytes):
    _header_end, _format_type, _track_count, chunks = (
        _parse_midi_chunks(midi_bytes)
    )
    track = next(chunk for chunk in chunks if chunk["id"] == b"MTrk")
    events, _end_tick = _parse_track_events(
        midi_bytes[track["data_start"]:track["data_end"]]
    )
    return [(tick, raw) for tick, _order, raw in events]


class FileInspectionTempoTests(unittest.TestCase):
    def test_all_midi_channels_have_distinct_legend_colors(self):
        colors = [
            _midi_channel_color(channel).name()
            for channel in range(1, 17)
        ]

        self.assertEqual(len(set(colors)), 16)
        self.assertEqual(_midi_channel_color(1).name(), colors[0])
        self.assertEqual(_midi_channel_color(16).name(), colors[-1])

    def test_tempo_percentage_accepts_very_slow_preview_values(self):
        self.assertEqual(_normalized_tempo_percent(5), 5)
        self.assertEqual(_normalized_tempo_percent(10), 10)
        self.assertEqual(_normalized_tempo_percent(1), 5)
        self.assertEqual(_normalized_tempo_percent(500), 400)

    def test_realtime_fluidsynth_uses_tempo_multiplier_commands(self):
        self.assertEqual(
            FluidSynthPlaybackProcess._tempo_command(5),
            "player_tempo_int 0.100001",
        )
        self.assertEqual(
            FluidSynthPlaybackProcess._tempo_command(10),
            "player_tempo_int 0.200000",
        )
        self.assertEqual(
            FluidSynthPlaybackProcess._tempo_command(175),
            "player_tempo_int 3.500000",
        )
        self.assertEqual(
            FluidSynthPlaybackProcess._tempo_command(400),
            "player_tempo_int 8.000000",
        )

    def test_realtime_fluidsynth_rebases_preview_tempo(self):
        source = _type_zero_midi(
            [
                (0, b"\xFF\x51\x03\x07\xA1\x20"),
                (96, bytes([0x90, 60, 100])),
            ]
        )

        rebased = FluidSynthPlaybackProcess._tempo_rebased_midi_bytes(
            source
        )
        tempo_events = [
            int.from_bytes(payload, "big")
            for _tick, raw in _track_events(rebased)
            for meta_type, payload in [_midi_meta_payload(raw)]
            if meta_type == 0x51
        ]

        self.assertEqual(tempo_events, [1000000])

    def test_tempo_percentage_scales_every_recorded_tempo(self):
        source = _type_zero_midi(
            [
                (0, b"\xFF\x51\x03\x07\xA1\x20"),
                (96, b"\xFF\x51\x03\x0F\x42\x40"),
                (96, bytes([0x90, 60, 100])),
                (192, bytes([0x80, 60, 0])),
            ]
        )

        faster = _scale_midi_tempo_bytes(source, 200)
        tempo_events = [
            (tick, int.from_bytes(payload, "big"))
            for tick, raw in _track_events(faster)
            for meta_type, payload in [_midi_meta_payload(raw)]
            if meta_type == 0x51
        ]
        timed_channel_events = _midi_output_events(faster)

        self.assertEqual(tempo_events, [(0, 250000), (96, 500000)])
        self.assertEqual(
            [raw for _seconds, raw in timed_channel_events],
            [bytes([0x90, 60, 100]), bytes([0x80, 60, 0])],
        )
        self.assertAlmostEqual(timed_channel_events[0][0], 0.25)
        self.assertAlmostEqual(timed_channel_events[1][0], 0.75)

    def test_source_seconds_convert_to_ticks_across_tempo_changes(self):
        source = _type_zero_midi(
            [
                (0, b"\xFF\x51\x03\x07\xA1\x20"),
                (96, b"\xFF\x51\x03\x0F\x42\x40"),
                (192, bytes([0x90, 60, 100])),
            ]
        )

        self.assertEqual(_midi_tick_for_seconds(source, 0.25), 48)
        self.assertEqual(_midi_tick_for_seconds(source, 0.75), 120)

    def test_tempo_percentage_scales_the_default_midi_tempo(self):
        source = _type_zero_midi(
            [
                (96, bytes([0x90, 64, 90])),
                (192, bytes([0x80, 64, 0])),
            ]
        )

        slower = _scale_midi_tempo_bytes(source, 50)
        tempo_events = [
            (tick, int.from_bytes(payload, "big"))
            for tick, raw in _track_events(slower)
            for meta_type, payload in [_midi_meta_payload(raw)]
            if meta_type == 0x51
        ]
        timed_channel_events = _midi_output_events(slower)

        self.assertEqual(tempo_events, [(0, 1000000)])
        self.assertAlmostEqual(timed_channel_events[0][0], 1.0)
        self.assertAlmostEqual(timed_channel_events[1][0], 2.0)

    def test_visual_timing_is_scaled_without_mutating_source_notes(self):
        notes = [{"start_sec": 2.0, "end_sec": 5.0, "pitch": 60}]

        scaled = _scale_preview_timed_items(notes, 200)

        self.assertEqual(scaled[0]["start_sec"], 1.0)
        self.assertEqual(scaled[0]["end_sec"], 2.5)
        self.assertEqual(notes[0]["start_sec"], 2.0)
        self.assertEqual(notes[0]["end_sec"], 5.0)

    def test_midi_output_worker_accepts_live_tempo_changes(self):
        class _Output:
            def send_message(self, _message):
                pass

        worker = MidiOutputWorker(b"", 0, tempo_percent=100)
        worker.set_tempo_percent(10)

        tempo = worker._drain_commands(
            _Output(),
            100,
            {},
            {},
        )

        self.assertEqual(tempo, 10)


if __name__ == "__main__":
    unittest.main()
