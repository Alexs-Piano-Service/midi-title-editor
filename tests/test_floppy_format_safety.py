import ctypes
from types import SimpleNamespace

import pytest

from aps_midi_prep_tool_app import main_window
from aps_midi_prep_tool_app.floppy_image import (
    DISK_FORMAT_BY_KEY,
    FloppyDriveInfo,
    FloppyImageError,
    FloppyImageSession,
    _WindowsVolumeHandle,
    usb_floppy_format_capacity_error,
)
from aps_midi_prep_tool_app.main_window import MidiTitleWindow


def _usb_drive(size_bytes):
    return FloppyDriveInfo(
        path="A:",
        size_bytes=size_bytes,
        transport="usb",
        model="Windows removable drive A:",
    )


def test_usb_format_capacity_rejects_2880_image_for_1440_drive():
    drive = _usb_drive(DISK_FORMAT_BY_KEY["ibm.1440"].size_bytes)
    disk_format = DISK_FORMAT_BY_KEY["ibm.2880"]

    message = usb_floppy_format_capacity_error(drive, disk_format)

    assert "IBM 2.88M ED" in message
    assert "2,949,120 bytes" in message
    assert "1,474,560 bytes" in message
    assert "will not attempt an oversized raw write" in message


@pytest.mark.parametrize("size_bytes", [0, 737_280, 1_474_560])
def test_usb_format_capacity_allows_720_image_when_capacity_is_sufficient_or_unknown(
    size_bytes,
):
    assert not usb_floppy_format_capacity_error(
        _usb_drive(size_bytes),
        DISK_FORMAT_BY_KEY["ibm.720"],
    )


def test_usb_format_session_blocks_oversized_write_before_disk_probe(monkeypatch):
    def unexpected_probe(*_args, **_kwargs):
        pytest.fail("The disk must not be probed or mutated after capacity validation fails")

    monkeypatch.setattr(
        FloppyImageSession,
        "_try_prepare_existing_usb_floppy",
        classmethod(unexpected_probe),
    )

    with pytest.raises(FloppyImageError, match="oversized raw write"):
        FloppyImageSession.format_usb_floppy(
            _usb_drive(DISK_FORMAT_BY_KEY["ibm.1440"].size_bytes),
            DISK_FORMAT_BY_KEY["ibm.2880"],
        )


def test_format_dialog_blocks_oversized_usb_write_before_confirmation(monkeypatch):
    drive = _usb_drive(DISK_FORMAT_BY_KEY["ibm.1440"].size_bytes)
    warning = {}

    class FakeWindow:
        def _prepare_for_disk_load(self, _description):
            return True

        def _choose_format_floppy_options(self):
            return {
                "source_kind": "floppy_usb",
                "source": drive,
                "target_name": drive.display_name,
                "drive_size_bytes": drive.size_bytes,
                "disk_format": DISK_FORMAT_BY_KEY["ibm.2880"],
                "eseq_disk": False,
            }

        def _confirm_format_floppy(self, *_args, **_kwargs):
            pytest.fail("An oversized format must be rejected before confirmation")

        def _start_floppy_format_worker(self, *_args, **_kwargs):
            pytest.fail("An oversized format must not start a worker")

    def capture_warning(_parent, title, message, *_args, **_kwargs):
        warning.update(title=title, message=message)

    monkeypatch.setattr(main_window.QMessageBox, "warning", capture_warning)

    MidiTitleWindow.format_disklavier_floppy(FakeWindow())

    assert warning["title"] == "Incompatible Floppy Format"
    assert "IBM 2.88M ED" in warning["message"]
    assert "will not attempt an oversized raw write" in warning["message"]


def test_windows_short_write_reports_image_and_chunk_byte_counts(tmp_path):
    image_path = tmp_path / "oversized.img"
    image_path.write_bytes(b"x" * 10_000)

    class FakeKernel32:
        def __init__(self):
            self.write_sizes = iter((8_192, 0))

        def WriteFile(self, _handle, _buffer, _size, bytes_written, _overlapped):
            bytes_written._obj.value = next(self.write_sizes)
            return True

        def FlushFileBuffers(self, _handle):
            pytest.fail("A short write must stop before flushing")

    volume = _WindowsVolumeHandle.__new__(_WindowsVolumeHandle)
    volume.path = r"\\.\A:"
    volume.handle = object()
    volume._ctypes = SimpleNamespace(
        byref=ctypes.byref,
        create_string_buffer=ctypes.create_string_buffer,
        FormatError=lambda _code: "",
        get_last_error=lambda: 0,
    )
    volume._wintypes = SimpleNamespace(DWORD=ctypes.c_uint32)
    volume._kernel32 = FakeKernel32()
    volume._seek = lambda _offset, _label: None

    with pytest.raises(FloppyImageError) as exc_info:
        volume.write_file(image_path)

    message = str(exc_info.value)
    assert r"Could not fully write floppy device \\.\A:." in message
    assert "Windows wrote 8,192 of 10,000 image bytes" in message
    assert "requested 1,808 bytes and wrote 0 bytes" in message
    assert "larger than the floppy or drive capacity" in message
