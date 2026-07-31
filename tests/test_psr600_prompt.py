import unittest

from aps_midi_prep_tool_app.main_window import (
    _psr600_conversion_prompt_copy,
)


class Psr600PromptTests(unittest.TestCase):
    def test_prompt_is_short_and_omits_extracted_filenames(self):
        headline, detail = _psr600_conversion_prompt_copy(
            {
                "file_count": 6,
                "melody_bank_count": 29,
                "apparent_layer_count": 1,
                "chord_bank_count": 10,
                "partial_melody_bank_count": 5,
                "filenames": [
                    "a76344fb48154a39b8b63b3edd60d9a3_PSR___01.BLK",
                ],
            },
            "DISK003.hfe",
        )

        self.assertEqual(
            headline,
            "Convert 6 PSR-600 Page Memory files to MIDI?",
        )
        self.assertEqual(
            detail,
            "DISK003.hfe contains 29 recorded Melody banks.\n\n"
            "Each BLK becomes one multitrack Type 1 MIDI using approximate "
            "General MIDI instruments. Source files are unchanged.\n\n"
            "• Conductor and Chord-bank switching are not decoded; "
            "banks may overlap.\n"
            "• 1 possible voice layer will be exported on a separate "
            "track/channel.\n"
            "• 10 Chord banks cannot be rendered.\n"
            "• 5 damaged Melody banks will stop at their last valid events.",
        )
        self.assertNotIn("PSR___01.BLK", detail)
        self.assertLess(len(detail), 500)

    def test_prompt_uses_singular_wording(self):
        headline, detail = _psr600_conversion_prompt_copy(
            {
                "file_count": 1,
                "melody_bank_count": 1,
                "partial_melody_bank_count": 1,
            }
        )

        self.assertEqual(
            headline,
            "Convert 1 PSR-600 Page Memory file to MIDI?",
        )
        self.assertIn("Found 1 recorded Melody bank.", detail)
        self.assertIn(
            "1 damaged Melody bank will stop at its last valid event.",
            detail,
        )


if __name__ == "__main__":
    unittest.main()
