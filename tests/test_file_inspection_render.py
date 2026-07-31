import os
import tempfile
import unittest
from unittest.mock import patch

from aps_midi_prep_tool_app.main_window import (
    BatchAudioRenderWorker,
    FileInspectionDialog,
    InspectionAudioRenderWorker,
    _midi_meta_payload,
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


class FileInspectionRenderTests(unittest.TestCase):
    def test_render_payload_applies_current_channel_instrument_and_tempo(self):
        source = _type_zero_midi(
            [
                (0, b"\xFF\x51\x03\x07\xA1\x20"),
                (0, bytes([0xC0, 0])),
                (0, bytes([0xC1, 5])),
                (24, bytes([0x90, 60, 100])),
                (24, bytes([0x91, 67, 100])),
                (48, bytes([0x80, 60, 0])),
                (48, bytes([0x81, 67, 0])),
            ]
        )

        class _Dialog:
            current_midi_bytes = source

            @staticmethod
            def _preview_channels():
                return {1}

            @staticmethod
            def _effective_channel_levels():
                return {1: 50}

            @staticmethod
            def _effective_channel_program_overrides():
                return {1: 40}

            @staticmethod
            def _preview_tempo_percent():
                return 200

            def _filtered_midi_bytes_for_preview(self, **kwargs):
                return FileInspectionDialog._filtered_midi_bytes_for_preview(
                    self,
                    **kwargs,
                )

        rendered_midi = FileInspectionDialog._inspection_render_midi_bytes(
            _Dialog()
        )
        events = _track_events(rendered_midi)

        tempo_values = [
            int.from_bytes(payload, "big")
            for _tick, raw in events
            for meta_type, payload in [_midi_meta_payload(raw)]
            if meta_type == 0x51
        ]
        channel_events = [
            raw for _tick, raw in events
            if raw and 0x80 <= raw[0] <= 0xEF
        ]

        self.assertEqual(tempo_values, [250000])
        self.assertIn(bytes([0xC0, 40]), channel_events)
        self.assertIn(bytes([0xB0, 7, 64]), channel_events)
        self.assertIn(bytes([0x90, 60, 100]), channel_events)
        self.assertNotIn(bytes([0xC1, 5]), channel_events)
        self.assertNotIn(bytes([0x91, 67, 100]), channel_events)

    def test_selected_song_worker_uses_preview_volume_and_replaces_output(
        self,
    ):
        calls = {}

        def _fake_render(
            midi_path,
            soundfont_path,
            wav_path,
            cancel_callback=None,
            gain=0.8,
        ):
            with open(midi_path, "rb") as handle:
                calls["midi"] = handle.read()
            calls["soundfont"] = soundfont_path
            calls["gain"] = gain
            calls["cancel_callback"] = cancel_callback
            with open(wav_path, "wb") as handle:
                handle.write(b"temporary wav")

        def _fake_convert(
            wav_path,
            output_path,
            output_format,
            cancel_callback=None,
        ):
            calls["format"] = output_format
            self.assertTrue(os.path.isfile(wav_path))
            self.assertIsNotNone(cancel_callback)
            with open(output_path, "wb") as handle:
                handle.write(b"rendered audio")

        with tempfile.TemporaryDirectory() as directory:
            output_path = os.path.join(directory, "selected.wav")
            with open(output_path, "wb") as handle:
                handle.write(b"old output")
            worker = InspectionAudioRenderWorker(
                b"prepared MIDI",
                "/fonts/selected.sf3",
                output_path,
                "wav",
                volume_percent=35,
            )

            with (
                patch(
                    "aps_midi_prep_tool_app.main_window."
                    "_render_midi_file_to_wav",
                    side_effect=_fake_render,
                ),
                patch(
                    "aps_midi_prep_tool_app.main_window."
                    "_convert_wav_for_audio_export",
                    side_effect=_fake_convert,
                ),
            ):
                worker.run()

            with open(output_path, "rb") as handle:
                output = handle.read()

        self.assertEqual(calls["midi"], b"prepared MIDI")
        self.assertEqual(calls["soundfont"], "/fonts/selected.sf3")
        self.assertAlmostEqual(calls["gain"], 0.28)
        self.assertEqual(calls["format"], "wav")
        self.assertEqual(output, b"rendered audio")

    def test_render_button_hands_current_song_and_settings_to_worker(self):
        calls = {}

        class _Signal:
            def connect(self, callback):
                calls.setdefault("connections", []).append(callback)

        class _Worker:
            def __init__(
                self,
                midi_bytes,
                soundfont_path,
                output_path,
                output_format,
                volume_percent=100,
                parent=None,
            ):
                calls.update(
                    {
                        "midi": midi_bytes,
                        "soundfont": soundfont_path,
                        "output": output_path,
                        "format": output_format,
                        "volume": volume_percent,
                        "parent": parent,
                    }
                )
                self.renderProgress = _Signal()
                self.renderFinished = _Signal()
                self.renderFailed = _Signal()
                self.finished = _Signal()

            def start(self):
                calls["started"] = True

        class _Combo:
            @staticmethod
            def currentData():
                return "/fonts/selected.sf3"

        class _Dialog:
            def __init__(self):
                self.inspection_render_worker = None
                self.soundfont_combo = _Combo()
                self.stopped = False
                self.rendering = False

            @staticmethod
            def t(text):
                return text

            @staticmethod
            def _suggested_inspection_render_path():
                return "/songs/selected.wav"

            @staticmethod
            def _inspection_render_midi_bytes():
                return b"current inspected song with settings"

            @staticmethod
            def _preview_volume_percent():
                return 42

            def _stop_playback(self):
                self.stopped = True

            def _set_inspection_rendering(self, rendering):
                self.rendering = rendering

            @staticmethod
            def _on_inspection_render_progress(*_args):
                pass

            @staticmethod
            def _on_inspection_render_finished(*_args):
                pass

            @staticmethod
            def _on_inspection_render_failed(*_args):
                pass

            @staticmethod
            def _on_inspection_render_worker_finished(*_args):
                pass

        dialog = _Dialog()
        with (
            patch(
                "aps_midi_prep_tool_app.main_window."
                "_find_fluidsynth_command",
                return_value="/usr/bin/fluidsynth",
            ),
            patch(
                "aps_midi_prep_tool_app.main_window.os.path.isfile",
                return_value=True,
            ),
            patch(
                "aps_midi_prep_tool_app.main_window."
                "QFileDialog.getSaveFileName",
                return_value=(
                    "/exports/selected.wav",
                    "WAV Audio (*.wav)",
                ),
            ),
            patch(
                "aps_midi_prep_tool_app.main_window."
                "InspectionAudioRenderWorker",
                _Worker,
            ),
        ):
            FileInspectionDialog._render_inspected_song(dialog)

        self.assertTrue(dialog.stopped)
        self.assertTrue(dialog.rendering)
        self.assertTrue(calls["started"])
        self.assertEqual(
            calls["midi"],
            b"current inspected song with settings",
        )
        self.assertEqual(calls["soundfont"], "/fonts/selected.sf3")
        self.assertEqual(calls["output"], "/exports/selected.wav")
        self.assertEqual(calls["format"], "wav")
        self.assertEqual(calls["volume"], 42)
        self.assertIs(calls["parent"], dialog)
        self.assertIs(dialog.inspection_render_worker.__class__, _Worker)

    def test_all_song_worker_applies_current_inspection_settings(self):
        source = _type_zero_midi(
            [
                (0, b"\xFF\x51\x03\x07\xA1\x20"),
                (0, bytes([0xC0, 0])),
                (0, bytes([0xC1, 5])),
                (24, bytes([0x90, 60, 100])),
                (24, bytes([0x91, 67, 100])),
                (48, bytes([0x80, 60, 0])),
                (48, bytes([0x81, 67, 0])),
            ]
        )
        rendered_payloads = []
        gains = []

        def _fake_render(
            midi_path,
            _soundfont_path,
            wav_path,
            cancel_callback=None,
            gain=0.8,
        ):
            with open(midi_path, "rb") as handle:
                rendered_payloads.append(handle.read())
            gains.append(gain)
            self.assertIsNotNone(cancel_callback)
            with open(wav_path, "wb") as handle:
                handle.write(b"temporary wav")

        def _fake_convert(
            _wav_path,
            output_path,
            _output_format,
            cancel_callback=None,
        ):
            self.assertIsNotNone(cancel_callback)
            with open(output_path, "wb") as handle:
                handle.write(b"rendered audio")

        with tempfile.TemporaryDirectory() as directory:
            items = []
            for name in ("one.mid", "two.mid"):
                path = os.path.join(directory, name)
                with open(path, "wb") as handle:
                    handle.write(source)
                items.append({"path": path, "label": name})

            worker = BatchAudioRenderWorker(
                items,
                "/fonts/selected.sf3",
                directory,
                "wav",
                channels={1},
                channel_levels={1: 50},
                program_overrides={1: 40},
                tempo_percent=200,
                volume_percent=25,
            )
            with (
                patch(
                    "aps_midi_prep_tool_app.main_window."
                    "_render_midi_file_to_wav",
                    side_effect=_fake_render,
                ),
                patch(
                    "aps_midi_prep_tool_app.main_window."
                    "_convert_wav_for_audio_export",
                    side_effect=_fake_convert,
                ),
            ):
                worker.run()

            self.assertTrue(
                os.path.isfile(os.path.join(directory, "one.wav"))
            )
            self.assertTrue(
                os.path.isfile(os.path.join(directory, "two.wav"))
            )

        self.assertEqual(len(rendered_payloads), 2)
        self.assertEqual(gains, [0.2, 0.2])
        for rendered_midi in rendered_payloads:
            events = _track_events(rendered_midi)
            tempo_values = [
                int.from_bytes(payload, "big")
                for _tick, raw in events
                for meta_type, payload in [_midi_meta_payload(raw)]
                if meta_type == 0x51
            ]
            channel_events = [
                raw
                for _tick, raw in events
                if raw and 0x80 <= raw[0] <= 0xEF
            ]
            self.assertEqual(tempo_values, [250000])
            self.assertIn(bytes([0xC0, 40]), channel_events)
            self.assertIn(bytes([0xB0, 7, 64]), channel_events)
            self.assertNotIn(bytes([0x91, 67, 100]), channel_events)

    def test_all_song_action_hands_all_items_and_settings_to_worker(
        self,
    ):
        calls = {}

        class _Signal:
            def connect(self, callback):
                calls.setdefault("connections", []).append(callback)

        class _Worker:
            def __init__(
                self,
                items,
                soundfont_path,
                output_dir,
                output_format,
                parent=None,
                **settings,
            ):
                calls.update(
                    {
                        "items": items,
                        "soundfont": soundfont_path,
                        "output_dir": output_dir,
                        "format": output_format,
                        "parent": parent,
                        "settings": settings,
                    }
                )
                self.renderProgress = _Signal()
                self.renderFinished = _Signal()
                self.renderFailed = _Signal()
                self.finished = _Signal()

            def start(self):
                calls["started"] = True

        class _Combo:
            @staticmethod
            def currentData():
                return "/fonts/selected.sf3"

        class _Dialog:
            def __init__(self):
                self.items = [
                    {"path": "/songs/one.mid"},
                    {"path": "/songs/two.mid"},
                ]
                self.inspection_render_worker = None
                self.soundfont_combo = _Combo()
                self.stopped = False
                self.rendering = False

            @staticmethod
            def t(text):
                return text

            @staticmethod
            def _suggested_inspection_render_directory():
                return "/songs"

            @staticmethod
            def _preview_channels():
                return {1, 3}

            @staticmethod
            def _effective_channel_levels():
                return {1: 45, 3: 80}

            @staticmethod
            def _effective_channel_program_overrides():
                return {1: 40}

            @staticmethod
            def _preview_tempo_percent():
                return 75

            @staticmethod
            def _preview_volume_percent():
                return 35

            def _stop_playback(self):
                self.stopped = True

            def _set_inspection_rendering(self, rendering):
                self.rendering = rendering

            @staticmethod
            def _on_inspection_batch_render_progress(*_args):
                pass

            @staticmethod
            def _on_inspection_batch_render_finished(*_args):
                pass

            @staticmethod
            def _on_inspection_render_failed(*_args):
                pass

            @staticmethod
            def _on_inspection_render_worker_finished(*_args):
                pass

        dialog = _Dialog()
        with (
            patch(
                "aps_midi_prep_tool_app.main_window."
                "_find_fluidsynth_command",
                return_value="/usr/bin/fluidsynth",
            ),
            patch(
                "aps_midi_prep_tool_app.main_window.os.path.isfile",
                return_value=True,
            ),
            patch(
                "aps_midi_prep_tool_app.main_window."
                "QFileDialog.getExistingDirectory",
                return_value="/exports",
            ),
            patch(
                "aps_midi_prep_tool_app.main_window."
                "BatchAudioRenderWorker",
                _Worker,
            ),
        ):
            FileInspectionDialog._render_all_inspected_songs(
                dialog,
                "wav",
            )

        self.assertTrue(dialog.stopped)
        self.assertTrue(dialog.rendering)
        self.assertTrue(calls["started"])
        self.assertEqual(calls["items"], dialog.items)
        self.assertEqual(calls["soundfont"], "/fonts/selected.sf3")
        self.assertEqual(calls["output_dir"], "/exports")
        self.assertEqual(calls["format"], "wav")
        self.assertIs(calls["parent"], dialog)
        self.assertEqual(calls["settings"]["channels"], {1, 3})
        self.assertEqual(
            calls["settings"]["channel_levels"],
            {1: 45, 3: 80},
        )
        self.assertEqual(
            calls["settings"]["program_overrides"],
            {1: 40},
        )
        self.assertEqual(calls["settings"]["tempo_percent"], 75)
        self.assertEqual(calls["settings"]["volume_percent"], 35)
        self.assertEqual(dialog._inspection_render_scope, "all")
        self.assertEqual(
            dialog._inspection_render_output_dir,
            "/exports",
        )


if __name__ == "__main__":
    unittest.main()
