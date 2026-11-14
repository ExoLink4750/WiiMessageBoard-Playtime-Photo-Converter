# WiiMessageBoard-Playtime-Photo-Converter
Application made to convert your Wii Message Board dumps from CDBackup homebrew into readable spreadsheets and PNGs via an interactive GUI

## What it can do:

### Read every dump-type from CDBackup homebrew

- Raw **vff dumps** (recommended - contains everything: play-time logs,
  messages, title IDs and photos)
- **cdb dumps** (`cdbackup.vff` or a cdb folder tree)
- Inputs can be a **folder**, a **.zip**, a **.rar** archive, or a bare
  **.vff** container
- **Two dumps can be scanned together** (for example from two consoles, or a
  vff plus a cdb of the same console) - they are merged automatically and any
  logs/messages/photos present in both are counted only once

### Convert data to a spreadsheet file (.xlsx)

- **Combined Totals** sheet plus **one sheet per day** of play-time history
- **Achievements** (medals, records, pro status, stamps, trophies)
- **Letters** (Animal Crossing City Folk picture posts, Nintendo mail, channel messages)
- **Personal notes** (hand-written board messages)
- **Game title + Title-ID matching** of games that posted on the message board (for vff dumps)
- Letter stationery (U8 archives) detection

### Browse everything before you convert

- **Titles tab**: every game/channel played, total play time and days played,
  sortable by any column
- **Days tab**: per-day play time, session counts and "Others" overflow flags
- **Memos tab**: every board message with category, game, message preview and a photo
  marker; double-click a row to read the complete message in a popup window, with any
  attached photos shown inline
- **Photos tab**: gallery of all attached photos, each with the date,
  the game/channel it came from and the full text of the message it accompanied

### Recover attached photos

- Decodes the Wii's proprietary **aJPG** photo format (a decoder is built in,
  no extra files needed)
- One-click export of all photos to PNG
- Photos embedded in the spreadsheet export are not supported - they are saved
  alongside it or packed into an "Export All" .zip together with the workbook


## Main project files

| File | Purpose |
|---|---|
| `playtime_log_gui.py` | The user interface (GUI, scanner, photo viewer, cache) |
| `playtime_log_to_excel.py` | Converter core - parsing and workbook building; also usable standalone from the command line |
| `ajpg_to_png.py` | AJPG photo decoder (optional - an embedded copy is built into the GUI as fallback) |


## Requirements for running from source + dependencies
  
Requires **Python 3.10+** with tkinter

```
pip install openpyxl pillow ttkbootstrap tkinterdnd2 rarfile
python playtime_log_gui.py
```

Only `openpyxl` and `pillow` are strictly required; the rest are optional
features (themed UI, drag and drop, .rar input) and can also be installed
later.

A **dependency manager** (Tools → Dependencies…) reports what is
installed and offers 1-click installs for missing packages.


### Command-line use for spreadsheet conversion (no GUI)

```
python playtime_log_to_excel.py <dump1> [dump2] -o "Playtime Log Summary.xlsx"
```

Accepts the same input types (folder path, .zip/.rar, .vff); two inputs are merged
with deduplication.

## How to quick launch:

With all three python scripts (playtime_log_to_excel.py, ajpg_to_png.py, and playtime_log_gui.py) in the same directory, run/double-click the "playtime_log_gui.py" script.

## Building the standalone executable on Windows

On Windows, place all three python scripts (playtime_log_to_excel.py, ajpg_to_png.py, and playtime_log_gui.py) in the same directory and build via CMD:
```
cd /d %~dp0

python -m pip install openpyxl pillow tkinterdnd2

:: Optional .rar support (rarfile).
python -m pip install rarfile
:: Optional themed UI + dark mode (ttkbootstrap).
python -m pip install ttkbootstrap

python -m pip install pyinstaller

python -m PyInstaller --noconfirm --onefile --windowed --name PlaytimeLogConverter --collect-all tkinterdnd2 --collect-all ttkbootstrap --hidden-import PIL._tkinter_finder --hidden-import ajpg_to_png --hidden-import playtime_log_to_excel playtime_log_gui.py

pause
```

It installs the dependencies and produces
a single-file `PlaytimeLogConverter.exe` with PyInstaller located in the "dist" directory where it was run.
