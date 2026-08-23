from aps_midi_prep_tool_app import disk_session_worker


def test_emulator_build_worker_forwards_include_subfolders_false(monkeypatch):
    calls = []
    result = object()

    def fake_build_emulator_disk_images(*args, **kwargs):
        calls.append((args, kwargs))
        return result

    monkeypatch.setattr(
        disk_session_worker,
        "build_emulator_disk_images",
        fake_build_emulator_disk_images,
    )
    worker = disk_session_worker.EmulatorImageBuildWorker(
        "source",
        "output",
        prefix="DSKA",
        starting_number=1,
        safety_margin_bytes=0,
        album_title="Album",
        disk_format=object(),
        output_ext="img",
        include_subfolders=False,
    )
    finished = []
    worker.buildFinished.connect(finished.append)

    worker.run()

    assert calls[0][1]["include_subfolders"] is False
    assert finished == [result]
