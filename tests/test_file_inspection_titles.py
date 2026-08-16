from aps_midi_prep_tool_app.main_window import FileInspectionDialog


def test_file_inspection_list_collapses_title_spaces_without_changing_filename():
    item = {
        "label": "My  File.mid -   Moon    River  ",
        "display_name": "My  File.mid",
        "title": "   Moon    River  ",
    }

    assert FileInspectionDialog._file_list_label(item) == "My  File.mid - Moon River"
    assert item["title"] == "   Moon    River  "


def test_file_inspection_list_omits_untitled_suffix():
    assert FileInspectionDialog._file_list_label(
        {
            "label": "SONG.MID",
            "display_name": "SONG.MID",
            "title": "Untitled",
        }
    ) == "SONG.MID"
