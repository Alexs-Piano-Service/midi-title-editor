from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QMessageBox, QCheckBox


def show_first_time_dialog():
    settings = QSettings("AlexPianoServiceLLC", "APSMidiTitleEditor")
    skip_dialog = settings.value("skip_first_time_dialog", False, type=bool)

    if not skip_dialog:
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setWindowTitle("Welcome to APS MIDI Title Editor")
        msgBox.setText("""<html>
      <head>
        <style type="text/css">
          body { font-family: "Helvetica", sans-serif; }
          h1 { font-weight: bold; margin-bottom: 10px; text-align: center; }
          h2 { text-align: center; }
          p { margin: 10px 0; }
          ul { margin: 10px 20px; }
          li { margin-bottom: 6px; }
          a { text-decoration: none; }
          a:hover { text-decoration: underline; }
        </style>
      </head>
      <body>
        <h1>APS MIDI Title Editor</h1>
        <p>
          This tool is designed to help you quickly update and manage the titles of your MIDI files.
          Whether you have a few or many files, renaming them is a breeze.
        </p>
        <p><strong>Getting Started:</strong></p>
        <ol>
          <li>
            <strong>Select a Folder:</strong> Click the <em>"Choose MIDI Folder"</em> button to load the folder containing your MIDI files.
          </li>
          <li>
            <strong>Or Drag and Drop:</strong> You can also drag <code>.mid</code> and <code>.midi</code> files directly into the table.
          </li>
          <li>
            <strong>Edit Titles:</strong> Click a title cell to open the edit dialog. Changes are queued until you save.
          </li>
          <li>
            <strong>Quick Copy:</strong> Click the clipboard icon or double-click a filename to copy the filename.
          </li>
          <li>
            <strong>Save Options:</strong> Use <em>Save</em>, <em>Save As...</em>, or <em>Rename All to DOS 8.3</em> from the Save button menu.
          </li>
        </ol>
        <p>
          For more details, visit our website at
          <a href="https://www.alexanderpeppe.com/">alexanderpeppe.com</a>.<br />
          Copyright 2025 Alex's Piano Service LLC
        </p>
        <p>Happy editing!</p>
      </body>
    </html>""")
        dont_show_checkbox = QCheckBox("Do not show this dialog again")
        dont_show_checkbox.setStyleSheet("margin-top: 10px")
        msgBox.setCheckBox(dont_show_checkbox)
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec()
        if dont_show_checkbox.isChecked():
            settings.setValue("skip_first_time_dialog", True)
