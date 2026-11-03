# VisionMacro Suite Pro

Windows keyboard/mouse macro recorder and player (customtkinter + pynput, optional OCR).

## Setup and run (Windows)

```bash
py -m pip install customtkinter pynput mss
py -m pip install easyocr numpy        # optional OCR
py vision_suite_pro_updated.py
```

## Data

- `vision_suite_save.json` holds environments, macros, hotkeys and variables. It is saved atomically, and the previous version is kept as `vision_suite_save.json.bak`.
- `vision_suite_run_log.txt` is a rotating error log (5 MB × 3). The current version logs only save errors; older versions also logged playback events, and those entries are still in the existing file.

To restore a bad save, close the app and copy `vision_suite_save.json.bak` over `vision_suite_save.json`.

## Safety

Python and OCR code blocks inside macros run with your Windows account's permissions, and only when `trustedScripts` is on. Only import macro files you trust.

## Tests

There are no automated tests: the app needs a Windows desktop session for keyboard and mouse hooks. `python -m py_compile vision_suite_pro_updated.py` checks syntax. To validate manually: record a short macro, play it once, restart the app, and confirm the macro is still there.

## Limitations

Screen coordinates are absolute, so macros recorded on one monitor layout may misfire on another.
