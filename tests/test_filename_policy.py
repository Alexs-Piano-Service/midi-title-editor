from types import MethodType, SimpleNamespace

import pytest

from aps_midi_prep_tool_app import floppy_image as floppy_image_module
from aps_midi_prep_tool_app.dos83_renamer import (
    build_dos83_filename,
    is_dos83_filename,
)
from aps_midi_prep_tool_app.floppy_image import (
    DISK_FORMATS,
    FloppyImageError,
    FloppyImageSession,
    create_floppy_images_from_files,
    read_image_listing,
)
from aps_midi_prep_tool_app.main_window import MidiTitleWindow, QMessageBox
from aps_midi_prep_tool_app.message_catalog import (
    COMMON_TEXT_TRANSLATIONS,
    SUPPORTED_LANGUAGES,
)


FILENAME_POLICY_TRANSLATION_SOURCES = (
    "Use 8.3 filenames",
    "Restrict newly entered and generated disk/image filenames to an eight-character name and three-character extension. Leave this off to preserve normal long filenames.",
    "Restrict the filename to an eight-character name and three-character extension. This choice is remembered in Settings.",
    "Rename File",
    "Invalid Filename",
    "Name Already Exists",
    "DOS 8.3 filename rules are enabled for new filename changes.",
    "Long filenames are enabled. Existing pending names were left unchanged.",
    "DOS 8.3 filenames must use printable ASCII characters only.",
    "DOS 8.3 filenames cannot contain spaces.",
    "Filename contains characters that are not valid in DOS 8.3 names.",
    "DOS 8.3 filenames can only contain one extension separator.",
    "DOS 8.3 filename base must be 8 characters or fewer.",
    "DOS 8.3 extension must be 3 characters or fewer.",
)


class FilenamePolicyWindow:
    IMAGE_FILENAME_INVALID_CHARS = MidiTitleWindow.IMAGE_FILENAME_INVALID_CHARS
    _language_code = MidiTitleWindow._language_code
    _lt = MidiTitleWindow._lt

    def __init__(self, *, use_dos83=False):
        self.use_dos83 = use_dos83

    def _dos83_filenames_enabled(self):
        return self.use_dos83

    def _normalize_image_filename(self, filename, *, enforce_dos83=None):
        return MidiTitleWindow._normalize_image_filename(
            self,
            filename,
            enforce_dos83=enforce_dos83,
        )

    def _validate_image_filename(self, filename, *, enforce_dos83=None):
        return MidiTitleWindow._validate_image_filename(
            self,
            filename,
            enforce_dos83=enforce_dos83,
        )

    def _build_dos_image_filename(self, filename, used_paths):
        return MidiTitleWindow._build_dos_image_filename(self, filename, used_paths)

    def _build_image_filename(self, filename, used_paths, *, enforce_dos83=None):
        return MidiTitleWindow._build_image_filename(
            self,
            filename,
            used_paths,
            enforce_dos83=enforce_dos83,
        )


class FakeItem:
    def __init__(self, text):
        self.value = text

    def text(self):
        return self.value

    def setText(self, value):
        self.value = value


class FakeTable:
    def __init__(self, paths):
        self.rows = [{1: FakeItem(path), 3: FakeItem(path)} for path in paths]

    def rowCount(self):
        return len(self.rows)

    def item(self, row, column):
        return self.rows[row].get(column)


class FakeStatusLabel:
    def __init__(self):
        self.value = ""

    def setText(self, value):
        self.value = value


class FakeSettings:
    def __init__(self, values=None):
        self.values = dict(values or {})

    def value(self, key, default=None, **_kwargs):
        return self.values.get(key, default)

    def contains(self, key):
        return key in self.values

    def setValue(self, key, value):
        self.values[key] = value


class FakeCheckableAction:
    def __init__(self, checked=False):
        self.checked = checked

    def isChecked(self):
        return self.checked

    def setChecked(self, checked):
        self.checked = bool(checked)


def test_long_filenames_are_allowed_by_default():
    window = FilenamePolicyWindow()

    assert window._normalize_image_filename("My Favorite Song.mid") == "My Favorite Song.mid"
    assert window._validate_image_filename("My Favorite Song.mid") is None


def test_dos83_checkbox_enforces_short_uppercase_names():
    window = FilenamePolicyWindow(use_dos83=True)

    assert window._normalize_image_filename("song.mid") == "SONG.MID"
    assert "8 characters or fewer" in window._validate_image_filename("ABCDEFGHI.MID")


def test_dos83_policy_is_backed_by_settings_action_instead_of_quick_panel():
    setting_key = MidiTitleWindow.SETTING_USE_DOS83_FILENAMES
    window = SimpleNamespace(
        SETTING_USE_DOS83_FILENAMES=setting_key,
        settings=FakeSettings({setting_key: False}),
        settingsUseDos83FilenamesAction=FakeCheckableAction(False),
        status_label=FakeStatusLabel(),
        _lt=lambda source: source,
    )

    assert not MidiTitleWindow._dos83_filenames_enabled(window)

    MidiTitleWindow.toggle_dos83_filenames(window, True)

    assert MidiTitleWindow._dos83_filenames_enabled(window)
    assert window.settingsUseDos83FilenamesAction.isChecked()
    assert "enabled" in window.status_label.value


def test_dos83_and_track_title_filename_preferences_are_inverse():
    window = SimpleNamespace(
        SETTING_USE_DOS83_FILENAMES=MidiTitleWindow.SETTING_USE_DOS83_FILENAMES,
        SETTING_LONG_MIDI_FILENAMES=MidiTitleWindow.SETTING_LONG_MIDI_FILENAMES,
        SETTING_ESEQ_TO_MIDI_LONG_FILENAMES=(
            MidiTitleWindow.SETTING_ESEQ_TO_MIDI_LONG_FILENAMES
        ),
        SETTING_READ_FLOPPY_LONG_FILENAMES=(
            MidiTitleWindow.SETTING_READ_FLOPPY_LONG_FILENAMES
        ),
        DEFAULT_LONG_MIDI_FILENAMES=MidiTitleWindow.DEFAULT_LONG_MIDI_FILENAMES,
        settings=FakeSettings(
            {
                MidiTitleWindow.SETTING_USE_DOS83_FILENAMES: False,
                MidiTitleWindow.SETTING_LONG_MIDI_FILENAMES: True,
                MidiTitleWindow.SETTING_ESEQ_TO_MIDI_LONG_FILENAMES: True,
                MidiTitleWindow.SETTING_READ_FLOPPY_LONG_FILENAMES: True,
            }
        ),
        settingsUseDos83FilenamesAction=FakeCheckableAction(False),
        status_label=FakeStatusLabel(),
        _lt=lambda source: source,
    )
    window._set_long_midi_filenames_enabled = MethodType(
        MidiTitleWindow._set_long_midi_filenames_enabled,
        window,
    )

    MidiTitleWindow.toggle_dos83_filenames(window, True)

    assert window.settings.values[window.SETTING_USE_DOS83_FILENAMES]
    assert not MidiTitleWindow._long_midi_filenames_enabled(window)
    assert not window.settings.values[window.SETTING_LONG_MIDI_FILENAMES]
    assert not window.settings.values[window.SETTING_ESEQ_TO_MIDI_LONG_FILENAMES]
    assert not window.settings.values[window.SETTING_READ_FLOPPY_LONG_FILENAMES]

    MidiTitleWindow._set_long_midi_filenames_enabled(window, True)

    assert MidiTitleWindow._long_midi_filenames_enabled(window)
    assert not window.settings.values[window.SETTING_USE_DOS83_FILENAMES]
    assert not window.settingsUseDos83FilenamesAction.isChecked()


def test_filename_policy_ui_has_complete_translation_coverage():
    translated_codes = {
        language.code
        for language in SUPPORTED_LANGUAGES
        if language.code != "en"
    }

    for source in FILENAME_POLICY_TRANSLATION_SOURCES:
        assert translated_codes <= set(COMMON_TEXT_TRANSLATIONS[source])


def test_image_name_builder_preserves_long_name_when_dos83_is_off():
    window = FilenamePolicyWindow()

    assert MidiTitleWindow._build_image_filename(
        window,
        "My Favorite Song.mid",
        set(),
    ) == "My Favorite Song.mid"


def test_eseq_image_conversion_forces_dos83_when_preference_is_off():
    window = FilenamePolicyWindow(use_dos83=False)
    window.imageEseqVariant = "disklavier"
    window._eseq_song_extension = lambda _variant=None: "FIL"
    window._image_used_filenames_for_directory = lambda *_args, **_kwargs: set()
    window._join_image_path = lambda directory, filename: (
        f"{directory}/{filename}" if directory else filename
    )

    target_path = MidiTitleWindow._converted_image_path_for_kind(
        window,
        "MUSIC/A Very Long Song Name.mid",
        "eseq",
    )

    target_name = target_path.rsplit("/", 1)[-1]
    stem, extension = target_name.rsplit(".", 1)
    assert len(stem) <= 8
    assert extension == "FIL"
    assert " " not in target_name


def test_existing_eseq_addition_forces_dos83_when_preference_is_off(monkeypatch):
    window = FilenamePolicyWindow(use_dos83=False)
    monkeypatch.setattr(
        "aps_midi_prep_tool_app.main_window.is_eseq_file",
        lambda _path: True,
    )

    target_name = MidiTitleWindow._build_image_addition_filename(
        window,
        "/songs/A Very Long Existing ESEQ Name.fil",
        set(),
    )

    stem, extension = target_name.rsplit(".", 1)
    assert len(stem) <= 8
    assert extension == "FIL"
    assert " " not in target_name


def test_save_preflight_stages_short_names_for_external_eseq_files():
    source_path = "/songs/A Very Long External ESEQ Name.fil"
    window = FilenamePolicyWindow(use_dos83=False)
    window.table = FakeTable([source_path])
    window.table.rows[0][3] = FakeItem("A Very Long External ESEQ Name.fil")
    window.is_image_mode = lambda: False
    window.is_local_eseq_mode = lambda: True
    window._regular_eseq_rows = lambda: (0,)
    window._regular_row_output_filename = MethodType(
        MidiTitleWindow._regular_row_output_filename,
        window,
    )
    window._regular_used_output_filenames_for_directory = (
        lambda *_args, **_kwargs: set()
    )
    staged = {}
    window._stage_regular_row_pending_rename = (
        lambda _row, source, target: staged.update({source: target})
    )

    MidiTitleWindow._ensure_eseq_filenames_dos83_for_save(window)

    target_name = staged[source_path]
    stem, extension = target_name.rsplit(".", 1)
    assert len(stem) <= 8
    assert extension == "FIL"
    assert " " not in target_name


def test_bulk_dos83_name_preserves_non_midi_extension():
    assert build_dos83_filename("A Very Long ESEQ Name.fil", 3).endswith(".FIL")
    assert build_dos83_filename("A Very Long ESEQ Name.fil", 3).startswith("03")


def test_dos83_validator_covers_eseq_disk_filename_rules():
    assert is_dos83_filename("PIANO001.FIL")
    assert is_dos83_filename("SONG.P01")
    assert not is_dos83_filename("A Very Long Song.FIL")
    assert not is_dos83_filename("TOO_LONG1.MDA")
    assert not is_dos83_filename("SONG NAME.FIL")
    assert not is_dos83_filename("SONG.")


def test_low_level_image_packer_rejects_long_eseq_name(tmp_path, monkeypatch):
    source_path = tmp_path / "staged-eseq.bin"
    source_path.write_bytes(b"E-SEQ fixture")
    monkeypatch.setattr(floppy_image_module, "is_eseq_file", lambda _path: True)

    with pytest.raises(FloppyImageError, match="DOS 8.3"):
        floppy_image_module._copy_host_file_into_image(
            str(tmp_path / "unused.img"),
            str(source_path),
            "A Very Long ESEQ Name.FIL",
        )


def test_bulk_dos83_utility_queues_image_renames(monkeypatch):
    paths = ["A Long MIDI Name.mid", "A Long ESEQ Name.fil"]

    class FakeWindow(FilenamePolicyWindow):
        def __init__(self):
            super().__init__()
            self.table = FakeTable(paths)
            self.image_session = SimpleNamespace(mode_name="Image Mode")
            self.pendingImageRenames = {}
            self.pendingImageAdditions = {}
            self.pendingImageExportFilenames = {}
            self.pendingImageTitleEdits = {}
            self.imageFileInfo = {}
            self.status_label = FakeStatusLabel()

        @staticmethod
        def _is_special_pianodir_row(_row):
            return False

        @staticmethod
        def _join_image_path(directory, filename):
            return f"{directory}/{filename}" if directory else filename

        def _final_image_path(self, source_path):
            return self.pendingImageRenames.get(source_path, source_path)

        def _refresh_image_filename_display(self, row):
            source_path = self.table.item(row, 1).text()
            self.table.item(row, 3).setText(
                self._final_image_path(source_path).split("/")[-1]
            )

        @staticmethod
        def _refresh_pianodir_row():
            return None

        @staticmethod
        def _refresh_image_mode_action_state():
            return None

        @staticmethod
        def _auto_fit_table_columns_after_batch_change():
            return None

    monkeypatch.setattr(
        QMessageBox,
        "question",
        staticmethod(lambda *_args, **_kwargs: QMessageBox.Yes),
    )
    window = FakeWindow()

    MidiTitleWindow._rename_all_image_files_dos83(window)

    assert set(window.pendingImageRenames) == set(paths)
    assert {name.rsplit(".", 1)[-1] for name in window.pendingImageRenames.values()} == {
        "MID",
        "FIL",
    }
    assert all(
        len(name.rsplit("/", 1)[-1].split(".", 1)[0]) <= 8
        for name in window.pendingImageRenames.values()
    )
    assert "Use Save or Save As Image" in window.status_label.value


def test_long_filename_round_trips_through_fat_image(tmp_path):
    source_path = tmp_path / "source.mid"
    source_path.write_bytes(b"long filename test")
    output_path = tmp_path / "long-name.img"
    disk_format = next(item for item in DISK_FORMATS if item.key == "ibm.1440")

    create_floppy_images_from_files(
        [
            {
                "host_path": str(source_path),
                "image_path": "My Favorite Song.mid",
                "display_name": "My Favorite Song.mid",
            }
        ],
        str(output_path),
        "img",
        disk_format,
    )

    listing = read_image_listing(str(output_path))
    assert [entry.name for entry in listing.entries] == ["My Favorite Song.mid"]

    session = FloppyImageSession.load(str(output_path))
    try:
        case_renamed_path = session.create_modified_image(
            renames={"My Favorite Song.mid": "my favorite song.mid"},
        )
        case_renamed_listing = read_image_listing(case_renamed_path)
        assert [entry.name for entry in case_renamed_listing.entries] == [
            "my favorite song.mid"
        ]

        renamed_path = session.create_modified_image(
            renames={"My Favorite Song.mid": "00MYFAVO.MID"},
        )
        renamed_listing = read_image_listing(renamed_path)
        assert [entry.name for entry in renamed_listing.entries] == ["00MYFAVO.MID"]
    finally:
        session.cleanup()
