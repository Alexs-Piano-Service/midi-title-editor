# APS MIDI Prep Tool

Preserve, convert, and organize music for Yamaha Disklavier and other
player-piano systems.

APS MIDI Prep Tool is a desktop app for reading old PianoSoft and E-SEQ
floppies, recovering songs, correcting titles, converting between Yamaha E-SEQ
and Standard MIDI, and preparing virtual disks for floppy emulators such as
Nalbantov. It brings these jobs together in one visual workflow so you can
review changes before exporting the result.

Current version: `0.8.1`

[Download the latest packaged release](https://github.com/Alexs-Piano-Service/aps-midi-prep-tool/releases/latest)
· [Read the guides](#guides-and-support) · [See what changed](CHANGELOG.md)

## What it helps you do

- **Preserve old piano disks.** Read physical floppies or disk images, attempt
  to recover songs from damaged media, and make IMG or SCP archives of physical
  disks before making changes.
- **Move music between systems.** Convert Yamaha E-SEQ to standard MIDI, MIDI
  back to E-SEQ, and MIDI Type 1 (SMF1) to the Type 0 (SMF0) format expected by
  many player pianos.
- **Prepare floppy-emulator libraries.** Turn a folder of songs into numbered
  IMG or HFE disks, with compatible filenames, song order, and Yamaha directory
  data created for each disk.
- **Clean up a music collection.** Correct song titles and filenames, remove
  Yamaha XF data, improve pedal compatibility, merge instruments for piano
  playback, or process a whole folder at once.
- **Check the result before saving.** Review metadata and piano-roll previews,
  audition songs with a SoundFont or USB MIDI device, and render WAV or MP3
  reference audio.

## How it works

1. Open a MIDI or E-SEQ folder, open a floppy image, or read a physical disk.
2. Review the songs and make any title, order, conversion, or compatibility
   changes you need.
3. Export the result to a folder, ZIP archive, new disk image, or physical
   floppy.

Edits and conversions made in the main file list are staged until you save.
Formatting and disk-writing commands run only after a separate confirmation,
and you can keep original images and floppies write-protected while exporting
clean copies.

## Popular workflows

| If you want to... | Start here |
| --- | --- |
| Back up a Yamaha floppy | Choose `Disk > Image Floppy...` for a direct IMG or SCP archive. To inspect the songs first, use `Disk > Read Floppy...`, then save the prepared working image or extract its files. If your operating system offers to format the disk, cancel. |
| Extract a collection of disk images | Choose `Utilities > Bulk Extraction...` to create a separate output folder for every supported image, with optional E-SEQ-to-MIDI conversion. |
| Build disks for a floppy emulator | Choose `Utilities > Build Emulator Disk Set...`, select a folder of songs, then create as many numbered HFE or IMG disks as needed. |
| Prepare MIDI for older hardware | Open the files, use the conversion and compatibility tools under `Utilities`, review the staged results, then save copies. |

The emulator-disk builder creates the disk image files themselves. Keep the
setup files from an existing Nalbantov USB stick and follow the emulator
manufacturer's instructions when preparing replacement USB media.

## Supported media and formats

| Source | What APS MIDI Prep Tool can do |
| --- | --- |
| Standard MIDI (`.mid`, `.midi`) | Edit, inspect, preview, batch-process, render, and convert to SMF0 or Yamaha E-SEQ. |
| Yamaha E-SEQ (`.FIL`/`.MDA` songs with `PIANODIR.FIL`/`MUSIC.DIR` catalogs) | Preserve album and song information, edit titles and order, convert to MIDI, and build compatible disks. |
| Physical 3.5-inch floppies | Read, image, format, recover, and write through a compatible floppy drive or Greaseweazle. |
| Floppy images | Open and create common IMG/BIN-style raw images and HFE images; create SCP flux captures from physical media. |
| Other legacy music formats | Extract or convert supported PianoDisc System 3, Akai MPC, Yamaha V50/SY77, Yamaha PSR-600, and Electone MDR data. |
| Exports | Save songs to a folder or ZIP, create IMG/HFE disk images, and render MIDI or E-SEQ to WAV or MP3. |

The main editing workflow is designed for Standard MIDI, Yamaha E-SEQ, and
Yamaha-compatible floppy media. Some additional legacy formats are import-only
or read-only conversion sources. Disklavier E-SEQ media uses DOS 8.3 filenames,
titles of up to 32 characters, and no more than 60 cataloged songs per disk.

## Designed for careful preservation

- Save an image of an irreplaceable floppy before editing or converting its
  contents. Recovery mode can retry difficult media, while Greaseweazle can
  show weak or unreadable disk areas and create a low-level SCP archive.
- Use `File > Write Protection > Write-Protect Original` to prevent an open
  image or floppy from being overwritten. `Save As`, `Save As Image`, and ZIP
  export remain available for copy-based work.
- Optional backups can be enabled under `File > Save Options`. Always test a
  newly written disk or image before relying on it.

## Requirements

Packaged releases are available for Windows and Linux. Source runs require
Python 3.10 or newer and PySide6. Depending on the workflow, you may also need:

- `mtools` for FAT floppy-image operations and the Greaseweazle CLI (`gw`) for
  Greaseweazle hardware, HFE conversion, and flux-image workflows.
- FluidSynth plus a compatible SoundFont for SoundFont-based previews and audio
  rendering, LAME for MP3 encoding, and `python-rtmidi` for direct MIDI-device
  playback. Basic piano preview does not require FluidSynth.
- Operating-system permission to access a physical floppy drive.

Availability of these optional tools varies by release package. FluidSynth and
a SoundFont are not bundled by default.

The interface is available in English, Spanish, French, German, Italian,
Brazilian Portuguese, Bulgarian, Dutch, Polish, Japanese, Korean, and
Simplified Chinese.

## Run from source

On Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install PySide6
python3 aps_midi_prep_tool.py
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for development and test guidance.

## Guides and support

- [Extract MIDI files from a Yamaha floppy disk](https://www.alexanderpeppe.com/extracting-midi-files-from-a-yamaha-floppy-disk-with-aps-midi-prep-tool/)
- [Change MIDI titles on your computer](https://www.alexanderpeppe.com/change-midi-titles-aps-midi-prep-tool/)
- [Copy a Yamaha PianoSoft floppy to a Nalbantov USB stick](https://www.alexanderpeppe.com/copying-a-yamaha-pianosoft-floppy-disk-to-a-nalbantov-usb-stick/)
- [Convert MIDI files from Type 1 to Type 0](https://www.alexanderpeppe.com/converting-midi-files-type-1-to-type-0-aps-midi-prep-tool/)
- [Edit titles and songs in Nalbantov virtual disks](https://www.alexanderpeppe.com/adding-removing-or-changing-titles-in-nalbantov-usb-stick-virtual-disks/)

Use `Help > Send Feedback...` or `Help > Report a Bug...` inside the app for
support. See [CHANGELOG.md](CHANGELOG.md) for release history and
[SECURITY.md](SECURITY.md) for responsible reporting guidance.

## License and responsible use

Copyright © 2026 Alex's Piano Service LLC. APS MIDI Prep Tool is released under
the [Apache License 2.0](LICENSE).

Use the tool only with disks and files you own or are authorized to preserve,
convert, or modify. Keep backups and test outputs before relying on them. The
project is an independent compatibility utility and is not affiliated with or
endorsed by Yamaha, Disklavier, PianoSoft, PianoDisc, Nalbantov, Greaseweazle,
Akai, or other companies and products named for compatibility purposes.

[Disclaimer](https://www.alexanderpeppe.com/disclaimer/) ·
[Privacy Policy](https://www.alexanderpeppe.com/privacy-policy/) ·
[DMCA Policy](https://www.alexanderpeppe.com/dmca-policy/)
