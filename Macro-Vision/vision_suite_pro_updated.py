"""VisionMacro Suite Pro 3.0
Python-only Windows macro recorder/player.

Required:
    python -m pip install customtkinter pynput mss
Optional OCR:
    python -m pip install easyocr numpy

Important: Python/OCR code blocks execute with your Windows account permissions.
Only run macro files and scripts you trust.
"""
from __future__ import annotations

import copy
import json
import logging
from logging.handlers import RotatingFileHandler
import multiprocessing as mp
import os
from pathlib import Path
import queue
import re
import threading
import time
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Any

import customtkinter as ctk
from pynput import keyboard, mouse
from pynput.keyboard import Key
from pynput.mouse import Button

APP_NAME = "VisionMacro Suite Pro"
VERSION = "3.3"
APP_DIR = Path(__file__).resolve().parent
SAVE_FILE = APP_DIR / "vision_suite_save.json"
LOG_FILE = APP_DIR / "vision_suite_run_log.txt"

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

logger = logging.getLogger("vision_suite")
logger.setLevel(logging.INFO)
handler = RotatingFileHandler(LOG_FILE, maxBytes=5_000_000, backupCount=3, encoding="utf-8")
handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S"))
logger.addHandler(handler)

KEY_ALIASES = {
    "ctrl_l": Key.ctrl, "ctrl_r": Key.ctrl, "ctrl": Key.ctrl,
    "shift_l": Key.shift, "shift_r": Key.shift, "shift": Key.shift,
    "alt_l": Key.alt, "alt_r": Key.alt, "alt": Key.alt,
    "enter": Key.enter, "tab": Key.tab, "esc": Key.esc, "space": Key.space,
    "backspace": Key.backspace, "delete": Key.delete, "up": Key.up,
    "down": Key.down, "left": Key.left, "right": Key.right,
    "home": Key.home, "end": Key.end, "page_up": Key.page_up,
    "page_down": Key.page_down, "cmd": Key.cmd, "win": Key.cmd,
}
# Windows OEM virtual-key codes can arrive without a character.
WINDOWS_OEM_VK_CHARS = {
    186: ";", 187: "=", 188: ",", 189: "-", 190: ".", 191: "/",
    192: "`", 219: "[", 220: "\\", 221: "]", 222: "'",
    109: "-", 107: "+", 106: "*", 111: "/", 110: ".",
}


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def key_to_string(value: Any) -> str:
    char = getattr(value, "char", None)
    if char is not None:
        if len(char) == 1 and 1 <= ord(char) <= 26:
            return chr(ord("a") + ord(char) - 1)
        return char
    vk = getattr(value, "vk", None)
    if vk in WINDOWS_OEM_VK_CHARS:
        return WINDOWS_OEM_VK_CHARS[vk]
    text = str(value)
    return {
        "Key.ctrl_l": "Key.ctrl", "Key.ctrl_r": "Key.ctrl",
        "Key.shift_l": "Key.shift", "Key.shift_r": "Key.shift",
        "Key.alt_l": "Key.alt", "Key.alt_r": "Key.alt",
    }.get(text, text)


def string_to_key(value: Any) -> Any:
    if isinstance(value, str):
        legacy_vk = re.fullmatch(r"<(\d+)>", value)
        if legacy_vk:
            vk = int(legacy_vk.group(1))
            return WINDOWS_OEM_VK_CHARS.get(vk, keyboard.KeyCode.from_vk(vk))
        if value.startswith("Key."):
            name = value.split(".", 1)[1]
            return KEY_ALIASES.get(name, getattr(Key, name, value))
    return value


def string_to_button(value: str) -> Any:
    return {
        "Button.left": Button.left,
        "Button.right": Button.right,
        "Button.middle": Button.middle,
        "Button.x1": getattr(Button, "x1", Button.left),
        "Button.x2": getattr(Button, "x2", Button.left),
    }.get(value, Button.left)


def blank_data() -> dict[str, Any]:
    return {
        "schemaVersion": 4,
        "current_env": "Default",
        "trustedScripts": True,
        "environments": {"Default": {"macros": {}, "variables": {}}},
    }


class Storage:
    @staticmethod
    def normalize(data: dict[str, Any]) -> None:
        data.setdefault("schemaVersion", 4)
        data["trustedScripts"] = True
        environments = data.setdefault("environments", {})
        if not environments:
            environments["Default"] = {"macros": {}, "variables": {}}
        if data.get("current_env") not in environments:
            data["current_env"] = next(iter(environments))
        for environment in environments.values():
            environment.setdefault("macros", {})
            environment.setdefault("variables", {})  # Legacy compatibility only; no UI.
            for macro in environment["macros"].values():
                macro.setdefault("events", [])
                macro.setdefault("hotkey", "None")
                macro.setdefault("speed", 1.0)
                macro.setdefault("loops", 1)

    def load(self) -> dict[str, Any]:
        if not SAVE_FILE.exists():
            return blank_data()
        data = json.loads(SAVE_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Save file root must be a JSON object.")
        self.normalize(data)
        return data

    def save(self, data: dict[str, Any]) -> None:
        self.normalize(data)
        payload = json.dumps(data, indent=2, ensure_ascii=False)
        temp = SAVE_FILE.with_suffix(".json.tmp")
        backup = SAVE_FILE.with_suffix(".json.bak")
        temp.write_text(payload, encoding="utf-8")
        json.loads(temp.read_text(encoding="utf-8"))
        if SAVE_FILE.exists():
            backup.write_bytes(SAVE_FILE.read_bytes())
        os.replace(temp, SAVE_FILE)


class Recorder:
    def __init__(self, callback) -> None:
        self.callback = callback
        self.active = False
        self.events: list[dict[str, Any]] = []
        self._last = 0.0
        self._keyboard_listener = None
        self._mouse_listener = None
        self._lock = threading.Lock()

    def start(self) -> None:
        if self.active:
            return
        self.events = []
        self.active = True
        self._last = time.monotonic()
        self._keyboard_listener = keyboard.Listener(on_press=self._press, on_release=self._release)
        self._mouse_listener = mouse.Listener(on_click=self._click, on_scroll=self._scroll)
        self._keyboard_listener.start()
        self._mouse_listener.start()
        self.callback("Recording. Press Escape to stop and keep captured events.")

    def stop(self) -> list[dict[str, Any]]:
        self.active = False
        for listener in (self._keyboard_listener, self._mouse_listener):
            if listener:
                try:
                    listener.stop()
                except Exception:
                    pass
        self._keyboard_listener = None
        self._mouse_listener = None
        return copy.deepcopy(self.events)

    def _append(self, event: dict[str, Any]) -> None:
        if not self.active:
            return
        now = time.monotonic()
        with self._lock:
            event["dt"] = round(min(max(now - self._last, 0), 10), 6)
            self._last = now
            self.events.append(event)

    def _press(self, key: Any):
        text = key_to_string(key)
        if text == "Key.esc":
            self.active = False
            self.callback("Recording stopped by Escape. Click Finish Recording to apply it.")
            return False
        self._append({"type": "key_down", "key": text})

    def _release(self, key: Any):
        text = key_to_string(key)
        if text != "Key.esc":
            self._append({"type": "key_up", "key": text})

    def _click(self, x: int, y: int, button: Any, pressed: bool) -> None:
        self._append({
            "type": "mouse_down" if pressed else "mouse_up",
            "x": x, "y": y, "button": str(button),
        })

    def _scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        self._append({"type": "mouse_scroll", "x": x, "y": y, "dx": dx, "dy": dy})


def release_inputs() -> None:
    """Best-effort cleanup in either the UI or playback process."""
    key_controller = keyboard.Controller()
    mouse_controller = mouse.Controller()
    for key in (Key.ctrl, Key.shift, Key.alt, Key.cmd):
        try:
            key_controller.release(key)
        except Exception:
            pass
    for button in (Button.left, Button.right, Button.middle):
        try:
            mouse_controller.release(button)
        except Exception:
            pass


def pixel_matches(capture, x: int, y: int, colour: str, tolerance: int) -> bool:
    raw = str(colour).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", raw):
        raise ValueError(f"Invalid colour {colour!r}")
    target = tuple(int(raw[i:i + 2], 16) for i in (0, 2, 4))
    blue, green, red = capture.grab({"left": int(x), "top": int(y), "width": 1, "height": 1}).pixel(0, 0)[:3]
    return all(abs(a - b) <= int(tolerance) for a, b in zip((red, green, blue), target))


def playback_process(payload: dict[str, Any], status_queue: mp.Queue) -> None:
    """Runs in a separate process so Escape can terminate even infinite scripts/OCR."""
    mouse_controller = mouse.Controller()
    key_controller = keyboard.Controller()
    capture = None
    held_keys: set[Any] = set()
    held_buttons: set[Any] = set()

    def status(text: str) -> None:
        try:
            status_queue.put_nowait(("status", text))
        except Exception:
            pass

    runtime_variables = copy.deepcopy(payload.get("legacyVariables", {}))
    def run_script(code: str, scope: dict[str, Any] | None = None) -> None:
        scope = dict(scope or {})
        scope["variables"] = runtime_variables
        for variable_name, variable_value in runtime_variables.items():
            if isinstance(variable_name, str) and variable_name.isidentifier():
                scope.setdefault(variable_name, variable_value)

        def type_text(text: Any) -> None:
            key_controller.type(str(text))

        def press(key: Any) -> None:
            resolved = string_to_key(key)
            key_controller.press(resolved)
            key_controller.release(resolved)

        def click(x: int, y: int, button: str = "Button.left") -> None:
            mouse_controller.position = (int(x), int(y))
            mouse_controller.click(string_to_button(button))

        scope.update({
            "type_text": type_text,
            "press": press,
            "click": click,
            "sleep": time.sleep,
            "time": time,
            "json": json,
            "re": re,
        })
        helpers = {"type_text", "press", "click", "sleep", "time", "json", "re", "variables"}
        exec(compile(code, "<vision-macro-block>", "exec"), {"__builtins__": __builtins__}, scope)
        for item_name, item_value in scope.items():
            if item_name not in helpers and item_name.isidentifier() and not item_name.startswith("__"):
                runtime_variables[item_name] = item_value

    def execute(event: dict[str, Any], macros: dict[str, Any], legacy_variables: dict[str, Any], depth: int) -> None:
        nonlocal capture
        kind = event.get("type")
        if kind in ("mouse_down", "mouse_up"):
            mouse_controller.position = (int(event["x"]), int(event["y"]))
            button = string_to_button(event.get("button", "Button.left"))
            if kind == "mouse_down":
                mouse_controller.press(button)
                held_buttons.add(button)
            else:
                mouse_controller.release(button)
                held_buttons.discard(button)
        elif kind == "mouse_click":
            mouse_controller.position = (int(event["x"]), int(event["y"]))
            mouse_controller.click(string_to_button(event.get("button", "Button.left")), int(event.get("count", 1)))
        elif kind == "mouse_move":
            mouse_controller.position = (int(event["x"]), int(event["y"]))
        elif kind == "mouse_scroll":
            mouse_controller.position = (int(event.get("x", 0)), int(event.get("y", 0)))
            mouse_controller.scroll(int(event.get("dx", 0)), int(event.get("dy", 0)))
        elif kind in ("key_down", "key_up"):
            key = string_to_key(event.get("key"))
            if kind == "key_down":
                key_controller.press(key)
                held_keys.add(key)
            else:
                key_controller.release(key)
                held_keys.discard(key)
        elif kind == "key_press":
            key = string_to_key(event.get("key", "Key.enter"))
            key_controller.press(key)
            key_controller.release(key)
        elif kind == "type_text":
            key_controller.type(str(event.get("text", "")))
        elif kind == "type_variable":
            # Compatibility for existing saves; variable editing was removed from the UI.
            name = str(event.get("variable", ""))
            if name not in legacy_variables:
                raise KeyError(f"Legacy variable {name!r} is missing")
            key_controller.type(str(legacy_variables[name]))
        elif kind == "wait":
            time.sleep(max(0, safe_float(event.get("seconds", 1), 1)))
        elif kind == "wait_condition":
            import mss
            if capture is None:
                capture = mss.mss()
            timeout = safe_float(event.get("timeout", 20), 20) or 20
            poll = max(0.05, safe_float(event.get("poll", 0.2), 0.2))
            deadline = time.monotonic() + timeout
            while time.monotonic() < deadline:
                if event.get("mode", "color") == "color" and pixel_matches(
                    capture, event["x"], event["y"], event["color"], event.get("tolerance", 8)
                ):
                    return
                time.sleep(poll)
            raise TimeoutError(f"Colour condition timed out after {timeout:g} seconds")
        elif kind == "play_macro":
            run_macro(str(event.get("macro_name", "")), macros, legacy_variables, depth + 1)
        elif kind == "run_python":
            run_script(str(event.get("code", "")))
        elif kind == "ocr_python":
            try:
                import easyocr
                import numpy as np
                import mss
            except ImportError as exc:
                raise RuntimeError("OCR block requires: python -m pip install easyocr numpy") from exc
            region = {
                "left": int(event.get("left", 0)), "top": int(event.get("top", 0)),
                "width": max(1, int(event.get("width", 300))),
                "height": max(1, int(event.get("height", 100))),
            }
            status("Loading OCR model...")
            with mss.mss() as ocr_capture:
                image = np.asarray(ocr_capture.grab(region))
            reader = easyocr.Reader(event.get("languages", ["en"]), gpu=False)
            results = reader.readtext(image, detail=1)
            ocr_text = " ".join(str(item[1]) for item in results)
            variable_name = str(event.get("variable_name", "ocr_text")).strip() or "ocr_text"
            if not variable_name.isidentifier():
                raise ValueError(f"Invalid OCR variable name: {variable_name!r}")
            runtime_variables[variable_name] = ocr_text
            status(f"OCR read into {variable_name}: {ocr_text[:80]}")
            run_script(str(event.get("code", "")), {
                variable_name: ocr_text,
                "ocr_text": ocr_text,
                "ocr_results": results,
                "region": region,
                "variables": legacy_variables,
            })
        else:
            raise ValueError(f"Unsupported event type {kind!r}")

    def run_macro(name: str, macros: dict[str, Any], legacy_variables: dict[str, Any], depth: int = 0) -> None:
        if depth > 12:
            raise RuntimeError("Nested macro depth exceeded 12")
        macro = macros.get(name)
        if not macro:
            raise KeyError(f"Macro {name!r} was not found")
        events = macro.get("events", [])
        speed = max(0.05, safe_float(macro.get("speed", 1), 1))
        loops = int(macro.get("loops", 1))
        loop_number = 0
        while loops == 0 or loop_number < max(1, loops):
            loop_number += 1
            label = "continuous" if loops == 0 else f"{loop_number}/{loops}"
            status(f"{name}: loop {label}")
            for index, event in enumerate(events, 1):
                time.sleep(max(0, safe_float(event.get("dt", 0), 0) / speed))
                status(f"{name}: loop {label}, step {index}/{len(events)}")
                execute(event, macros, legacy_variables, depth)

    try:
        run_macro(payload["macroName"], payload["macros"], payload.get("legacyVariables", {}))
        status_queue.put(("done", "Completed"))
    except Exception as exc:
        status_queue.put(("done", f"Failed: {exc}"))
    finally:
        for key in list(held_keys):
            try:
                key_controller.release(key)
            except Exception:
                pass
        for button in list(held_buttons):
            try:
                mouse_controller.release(button)
            except Exception:
                pass
        release_inputs()


class PlaybackManager:
    def __init__(self, ui_queue: queue.Queue) -> None:
        self.ui_queue = ui_queue
        self.process: mp.Process | None = None
        self.process_queue: mp.Queue | None = None
        self.monitor_thread: threading.Thread | None = None

    @property
    def playing(self) -> bool:
        return self.process is not None and self.process.is_alive()

    def start(self, payload: dict[str, Any]) -> None:
        if self.playing:
            raise RuntimeError("A macro is already running")
        self.process_queue = mp.Queue()
        self.process = mp.Process(target=playback_process, args=(payload, self.process_queue), daemon=True)
        self.process.start()
        self.monitor_thread = threading.Thread(target=self._monitor, daemon=True)
        self.monitor_thread.start()

    def _monitor(self) -> None:
        delivered_done = False
        while self.process is not None and self.process.is_alive():
            try:
                message = self.process_queue.get(timeout=0.1)
                if message[0] == "done":
                    delivered_done = True
                self.ui_queue.put(message)
            except queue.Empty:
                pass
            except (EOFError, OSError):
                break
        if self.process is not None:
            self.process.join(timeout=0.2)
        if not delivered_done:
            self.ui_queue.put(("done", "Stopped"))

    def kill(self) -> None:
        process = self.process
        if process is not None and process.is_alive():
            process.terminate()
            process.join(timeout=1.0)
            if process.is_alive() and hasattr(process, "kill"):
                process.kill()
                process.join(timeout=1.0)
        release_inputs()
        self.ui_queue.put(("done", "Stopped by Escape"))


EVENT_TEMPLATES = {
    "Mouse click": {"dt": 0.1, "type": "mouse_click", "x": 0, "y": 0, "button": "Button.left", "count": 1},
    "Key press": {"dt": 0.1, "type": "key_press", "key": "Key.enter"},
    "Type text": {"dt": 0.1, "type": "type_text", "text": ""},
    "Wait": {"dt": 0.1, "type": "wait", "seconds": 1.0},
    "Wait for colour": {"dt": 0.1, "type": "wait_condition", "mode": "color", "x": 0, "y": 0, "color": "#FFFFFF", "tolerance": 8, "timeout": 20, "poll": 0.2},
    "Play macro": {"dt": 0.1, "type": "play_macro", "macro_name": ""},
    "Python block": {"dt": 0.1, "type": "run_python", "code": "# Custom Python\n# type_text('hello')\n"},
    "OCR + Python block": {
        "dt": 0.1, "type": "ocr_python", "left": 0, "top": 0,
        "width": 400, "height": 160, "languages": ["en"], "variable_name": "ocr_text",
        "code": "# ocr_text contains the recognized text\nif 'example' in ocr_text.lower():\n    type_text('matched')\n",
    },
}


def event_summary(event: dict[str, Any]) -> str:
    kind = event.get("type", "?")
    if kind in ("mouse_down", "mouse_up", "mouse_click", "mouse_move"):
        return f"{kind}: ({event.get('x')}, {event.get('y')}) {event.get('button', '')}"
    if kind in ("key_down", "key_up", "key_press"):
        return f"{kind}: {event.get('key', '')}"
    if kind == "type_text":
        return f"type text: {str(event.get('text', ''))[:55]}"
    if kind == "type_variable":
        return f"legacy variable: {event.get('variable', '')}"
    if kind == "run_python":
        return "Python block"
    if kind == "ocr_python":
        return f"OCR + Python: {event.get('width')}x{event.get('height')} at ({event.get('left')}, {event.get('top')})"
    if kind == "play_macro":
        return f"play macro: {event.get('macro_name', '')}"
    if kind == "wait_condition":
        return f"wait for colour: timeout {event.get('timeout', 20)}s"
    if kind == "wait":
        return f"wait: {event.get('seconds', 1)}s"
    return str(kind)


class ScreenRegionSelector(ctk.CTkToplevel):
    """Transparent click-and-drag selector using virtual-screen coordinates."""
    def __init__(self, parent, pixel_mode: bool = False) -> None:
        super().__init__(parent)
        self.result = None
        self.pixel_mode = pixel_mode
        self.start = None
        self.overrideredirect(True)
        left, top = self.winfo_vrootx(), self.winfo_vrooty()
        width, height = self.winfo_vrootwidth(), self.winfo_vrootheight()
        self.geometry(f"{width}x{height}{left:+d}{top:+d}")
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.28)
        self.configure(fg_color="#1264a3")
        self.canvas = tk.Canvas(self, bg="#1264a3", highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)
        prompt = "Click the pixel to lock its colour" if pixel_mode else "Drag over the OCR area, then release to lock it in"
        self.canvas.create_text(width // 2, 45, text=prompt + "   |   Esc cancels", fill="white", font=("Segoe UI", 18, "bold"))
        self.box = None
        self.bind("<Escape>", lambda _e: self.destroy())
        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._drag)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        self.focus_force()
    def _press(self, event):
        self.start = (event.x, event.y)
        if self.pixel_mode:
            self._finish_pixel(event)
        else:
            self.box = self.canvas.create_rectangle(event.x, event.y, event.x, event.y, outline="white", width=3)
    def _drag(self, event):
        if self.box and self.start:
            self.canvas.coords(self.box, self.start[0], self.start[1], event.x, event.y)
    def _release(self, event):
        if self.pixel_mode or not self.start:
            return
        x1, y1 = self.start
        x2, y2 = event.x, event.y
        left, top = min(x1, x2) + self.winfo_vrootx(), min(y1, y2) + self.winfo_vrooty()
        width, height = max(1, abs(x2 - x1)), max(1, abs(y2 - y1))
        self.result = (left, top, width, height)
        self.destroy()
    def _finish_pixel(self, event):
        x, y = event.x + self.winfo_vrootx(), event.y + self.winfo_vrooty()
        self.withdraw()
        self.update_idletasks()
        time.sleep(0.08)
        try:
            import mss
            with mss.mss() as capture:
                blue, green, red = capture.grab({"left": x, "top": y, "width": 1, "height": 1}).pixel(0, 0)[:3]
            self.result = (x, y, f"#{red:02X}{green:02X}{blue:02X}")
        finally:
            self.destroy()

class EventDialog(ctk.CTkToplevel):
    KEYWORDS = set("False None True and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield".split())
    def __init__(self, parent, event: dict[str, Any], title: str = "Edit event") -> None:
        super().__init__(parent)
        self.result = None
        self.event = copy.deepcopy(event)
        self.kind = str(self.event.get("type", ""))
        self.title(title)
        self.geometry("820x680")
        self.minsize(680, 520)
        self.transient(parent)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(self, text=f"{title}: {event_summary(self.event)}", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", padx=16, pady=12)
        self.fields = ctk.CTkFrame(self)
        self.fields.grid(row=1, column=0, sticky="ew", padx=16)
        self.inputs = {}
        if self.kind == "ocr_python":
            for label, key in (("X", "left"), ("Y", "top"), ("Width", "width"), ("Height", "height"), ("Output variable", "variable_name")):
                self._field(label, key)
            ctk.CTkButton(self.fields, text="Select OCR region", command=self.select_region).pack(side="left", padx=6, pady=8)
        elif self.kind == "wait_condition":
            for label, key in (("X", "x"), ("Y", "y"), ("Tolerance", "tolerance"), ("Timeout", "timeout"), ("Poll", "poll")):
                self._field(label, key)
            self._field("Colour", "color")
            self.preview = ctk.CTkLabel(self.fields, text="   ", width=40, corner_radius=5, fg_color=self.event.get("color", "#FFFFFF"))
            self.preview.pack(side="left", padx=5)
            ctk.CTkButton(self.fields, text="Pick and lock pixel", command=self.select_pixel).pack(side="left", padx=6, pady=8)
        if self.kind in ("run_python", "ocr_python"):
            self.editor = tk.Text(self, undo=True, wrap="none", bg="#171717", fg="#e8e8e8", insertbackground="white", font=("Consolas", 12), relief="flat")
            self.editor.grid(row=2, column=0, sticky="nsew", padx=16, pady=8)
            self.editor.insert("1.0", str(self.event.get("code", "")))
            self.editor.tag_configure("keyword", foreground="#569CD6")
            self.editor.tag_configure("string", foreground="#CE9178")
            self.editor.tag_configure("comment", foreground="#6A9955")
            self.editor.bind("<KeyRelease>", self._highlight)
            self.validation = tk.StringVar(value="Python syntax: checking...")
            ctk.CTkLabel(self, textvariable=self.validation, anchor="w").grid(row=3, column=0, sticky="ew", padx=16)
            self.after(30, self._highlight)
        else:
            self.editor = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Consolas", size=13), undo=True)
            self.editor.grid(row=2, column=0, sticky="nsew", padx=16, pady=8)
            self.editor.insert("1.0", json.dumps(self.event, indent=2, ensure_ascii=False))
        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.grid(row=4, column=0, pady=12)
        ctk.CTkButton(buttons, text="Apply", command=self.apply).pack(side="left", padx=5)
        ctk.CTkButton(buttons, text="Cancel", fg_color="#666666", command=self.destroy).pack(side="left", padx=5)
    def _field(self, label, key):
        frame = ctk.CTkFrame(self.fields, fg_color="transparent")
        frame.pack(side="left", padx=4, pady=5)
        ctk.CTkLabel(frame, text=label).pack()
        value = tk.StringVar(value=str(self.event.get(key, "")))
        self.inputs[key] = value
        ctk.CTkEntry(frame, textvariable=value, width=88).pack()
    def select_region(self):
        self.grab_release(); self.withdraw()
        selector = ScreenRegionSelector(self.master)
        self.wait_window(selector); self.deiconify(); self.grab_set()
        if selector.result:
            for key, value in zip(("left", "top", "width", "height"), selector.result): self.inputs[key].set(str(value))
    def select_pixel(self):
        self.grab_release(); self.withdraw()
        selector = ScreenRegionSelector(self.master, pixel_mode=True)
        self.wait_window(selector); self.deiconify(); self.grab_set()
        if selector.result:
            x, y, colour = selector.result
            self.inputs["x"].set(str(x)); self.inputs["y"].set(str(y)); self.inputs["color"].set(colour)
            self.preview.configure(fg_color=colour)
    def _highlight(self, _event=None):
        code = self.editor.get("1.0", "end-1c")
        for tag in ("keyword", "string", "comment"): self.editor.tag_remove(tag, "1.0", "end")
        for match in re.finditer(r"#[^\n]*", code): self._tag("comment", match.start(), match.end(), code)
        for match in re.finditer(r"(?s)(['\"]{3}.*?['\"]{3}|'(?:\\.|[^'\\])*'|\"(?:\\.|[^\"\\])*\")", code): self._tag("string", match.start(), match.end(), code)
        for match in re.finditer(r"\b[A-Za-z_]\w*\b", code):
            if match.group() in self.KEYWORDS: self._tag("keyword", match.start(), match.end(), code)
        try:
            compile(code, "<macro-editor>", "exec")
            self.validation.set("Python syntax: valid")
        except SyntaxError as exc:
            self.validation.set(f"Syntax error, line {exc.lineno}: {exc.msg}")
    def _tag(self, tag, start, end, code):
        line = code.count("\n", 0, start) + 1; last = code.rfind("\n", 0, start)
        a = f"{line}.{start - last - 1}"
        line2 = code.count("\n", 0, end) + 1; last2 = code.rfind("\n", 0, end)
        b = f"{line2}.{end - last2 - 1}"
        self.editor.tag_add(tag, a, b)
    def apply(self):
        try:
            if self.kind in ("run_python", "ocr_python"):
                code = self.editor.get("1.0", "end-1c")
                compile(code, "<macro-editor>", "exec")
                value = copy.deepcopy(self.event); value["code"] = code
                for key, var in self.inputs.items():
                    raw = var.get().strip()
                    value[key] = raw if key == "variable_name" else int(raw)
                if self.kind == "ocr_python" and not value["variable_name"].isidentifier(): raise ValueError("Output variable must be a valid Python name")
            elif self.kind == "wait_condition":
                value = copy.deepcopy(self.event)
                for key, var in self.inputs.items():
                    raw = var.get().strip()
                    value[key] = raw if key == "color" else (float(raw) if key in ("timeout", "poll") else int(raw))
                if not re.fullmatch(r"#[0-9A-Fa-f]{6}", value["color"]): raise ValueError("Colour must look like #RRGGBB")
            else:
                value = json.loads(self.editor.get("1.0", "end").strip())
            if not isinstance(value, dict) or not value.get("type"): raise ValueError("Event must be an object with a type")
            value["dt"] = max(0, safe_float(value.get("dt", 0.1), 0.1)); self.result = value; self.destroy()
        except Exception as exc: messagebox.showerror("Invalid event", str(exc), parent=self)

class HotkeyCaptureDialog(ctk.CTkToplevel):
    def __init__(self, parent, current: str) -> None:
        super().__init__(parent)
        self.result = None
        self.keys: set[str] = set()
        self.title("Record custom hotkey")
        self.geometry("460x250")
        self.transient(parent)
        self.grab_set()
        self.label = ctk.CTkLabel(self, text="Press the desired combination\nExample: Ctrl + Alt + H", font=ctk.CTkFont(size=17))
        self.label.pack(expand=True, padx=20, pady=20)
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=12)
        ctk.CTkButton(button_frame, text="Clear hotkey", command=self.clear).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Cancel", fg_color="#666666", command=self.destroy).pack(side="left", padx=5)
        self.listener = keyboard.Listener(on_press=self.press, on_release=self.release)
        self.listener.start()
        self.protocol("WM_DELETE_WINDOW", self.close)

    def press(self, key: Any) -> None:
        text = key_to_string(key)
        if text == "Key.esc":
            return
        normalized = text.replace("Key.", "")
        self.keys.add(normalized)
        self.after(0, lambda: self.label.configure(text=" + ".join(sorted(self.keys))))

    def release(self, key: Any) -> None:
        if not self.keys:
            return
        modifiers = {"ctrl", "shift", "alt", "cmd"}
        non_modifiers = [item for item in self.keys if item not in modifiers]
        if non_modifiers:
            ordered = [item for item in ("ctrl", "shift", "alt", "cmd") if item in self.keys] + sorted(non_modifiers)
            self.result = "+".join(f"<{item}>" if item in modifiers or len(item) > 1 else item for item in ordered)
            self.after(0, self.close)

    def clear(self) -> None:
        self.result = "None"
        self.close()

    def close(self) -> None:
        try:
            self.listener.stop()
        except Exception:
            pass
        self.destroy()


class VisionSuiteApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_NAME} {VERSION}")
        self.geometry("1320x800")
        self.minsize(1000, 650)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.storage = Storage()
        try:
            self.data = self.storage.load()
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not load save file:\n{exc}")
            self.data = blank_data()
        self.ui_queue: queue.Queue = queue.Queue()
        self.playback = PlaybackManager(self.ui_queue)
        self.recorder = Recorder(self.report)
        self.global_hotkeys = None
        self._selected_macro_name = None
        self._selected_event_iid = None
        self._autosave_job = None
        self.escape_listener = keyboard.Listener(on_press=self._global_key_press)
        self.escape_listener.start()
        self._build_ui()
        self.refresh_all()
        self.after(75, self.drain_queue)

    @property
    def environment(self) -> dict[str, Any]:
        return self.data["environments"][self.data["current_env"]]

    def _build_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        header = ctk.CTkFrame(self, corner_radius=0)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        ctk.CTkLabel(header, text=APP_NAME, font=ctk.CTkFont(size=22, weight="bold")).pack(side="left", padx=18, pady=12)
        ctk.CTkLabel(header, text="ESC = emergency stop", text_color="#ff8080").pack(side="left", padx=12)
        self.environment_menu = ctk.CTkOptionMenu(header, command=self.change_environment)
        self.environment_menu.pack(side="right", padx=10)
        ctk.CTkButton(header, text="+ Environment", width=110, command=self.add_environment).pack(side="right", padx=3)

        sidebar = ctk.CTkFrame(self, width=280, corner_radius=0)
        sidebar.grid(row=1, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        ctk.CTkLabel(sidebar, text="Macros", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=14, pady=(14, 6))
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *_: self.refresh_macros())
        ctk.CTkEntry(sidebar, textvariable=self.search, placeholder_text="Search macros...").pack(fill="x", padx=12)
        self.macro_list = tk.Listbox(sidebar, bg="#202020", fg="white", selectbackground="#2878b5", relief="flat", font=("Segoe UI", 11), exportselection=False)
        self.macro_list.pack(fill="both", expand=True, padx=12, pady=10)
        self.macro_list.bind("<<ListboxSelect>>", self._on_macro_selected)
        macro_buttons = ctk.CTkFrame(sidebar, fg_color="transparent")
        macro_buttons.pack(fill="x", padx=8, pady=8)
        ctk.CTkButton(macro_buttons, text="New", width=58, command=self.new_macro).pack(side="left", padx=2)
        ctk.CTkButton(macro_buttons, text="Rename", width=68, command=self.rename_macro).pack(side="left", padx=2)
        ctk.CTkButton(macro_buttons, text="Delete", width=62, fg_color="#993535", command=self.delete_macro).pack(side="left", padx=2)
        ctk.CTkButton(macro_buttons, text="Copy", width=58, command=self.copy_macro).pack(side="left", padx=2)

        main = ctk.CTkFrame(self, corner_radius=0)
        main.grid(row=1, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)
        self.macro_title = ctk.CTkLabel(main, text="Select a macro", font=ctk.CTkFont(size=20, weight="bold"))
        self.macro_title.grid(row=0, column=0, sticky="w", padx=18, pady=(14, 4))

        controls = ctk.CTkFrame(main, fg_color="transparent")
        controls.grid(row=1, column=0, sticky="ew", padx=14, pady=4)
        ctk.CTkButton(controls, text="Play", width=90, fg_color="#278443", command=self.play).pack(side="left", padx=3)
        ctk.CTkButton(controls, text="STOP (ESC)", width=105, fg_color="#b33434", command=self.emergency_stop).pack(side="left", padx=3)
        ctk.CTkButton(controls, text="Record", width=90, command=self.start_recording).pack(side="left", padx=3)
        ctk.CTkButton(controls, text="Finish Recording", width=125, command=self.finish_recording).pack(side="left", padx=3)
        ctk.CTkButton(controls, text="Hotkey", width=90, command=self.edit_hotkey).pack(side="left", padx=3)
        ctk.CTkButton(controls, text="Loop / Speed", width=105, command=self.edit_playback_settings).pack(side="left", padx=3)
        ctk.CTkButton(controls, text="Reset Hotkey", width=105, command=self.reset_hotkey).pack(side="left", padx=3)
        ctk.CTkLabel(controls, text="Auto-save on", text_color="#75d18b").pack(side="right", padx=8)

        event_tools = ctk.CTkFrame(main)
        event_tools.grid(row=2, column=0, sticky="ew", padx=16, pady=(7, 3))
        ctk.CTkButton(event_tools, text="+ Event", width=85, command=self.add_event).pack(side="left", padx=3, pady=7)
        ctk.CTkButton(event_tools, text="+ Python", width=90, fg_color="#6c43a8", command=lambda: self.add_specific("Python block")).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="+ OCR/Python", width=110, fg_color="#6c43a8", command=lambda: self.add_specific("OCR + Python block")).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="+ Color Check", width=105, fg_color="#a8682a", command=lambda: self.add_specific("Wait for colour")).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="Edit", width=75, command=self.edit_event).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="Duplicate", width=85, command=self.duplicate_event).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="Delete", width=75, fg_color="#993535", command=self.delete_event).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="Move Up", width=85, command=lambda: self.move_event(-1)).pack(side="left", padx=3)
        ctk.CTkButton(event_tools, text="Move Down", width=90, command=lambda: self.move_event(1)).pack(side="left", padx=3)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#1d1d1d", foreground="white", fieldbackground="#1d1d1d", rowheight=27)
        style.map("Treeview", background=[("selected", "#2878b5")])
        self.event_tree = ttk.Treeview(main, columns=("number", "delay", "action"), show="headings", selectmode="extended", takefocus=True)
        self.event_tree.heading("number", text="#")
        self.event_tree.heading("delay", text="Delay")
        self.event_tree.heading("action", text="Action")
        self.event_tree.column("number", width=50, anchor="center", stretch=False)
        self.event_tree.column("delay", width=90, anchor="center", stretch=False)
        self.event_tree.column("action", width=700)
        self.event_tree.grid(row=3, column=0, sticky="nsew", padx=16, pady=4)
        self.event_tree.bind("<Double-1>", lambda _e: self.edit_event())

        self.status = tk.StringVar(value="Ready")
        ctk.CTkLabel(main, textvariable=self.status, anchor="w").grid(row=4, column=0, sticky="ew", padx=16, pady=8)

    def _on_macro_selected(self, _event=None) -> None:
        selected = self.macro_list.curselection()
        if not selected:
            return
        name = str(self.macro_list.get(selected[0]))
        if name != self._selected_macro_name:
            self._selected_event_iid = None
        self._selected_macro_name = name
        self.show_macro()

    def selected_macro(self) -> str | None:
        selected = self.macro_list.curselection()
        if selected:
            self._selected_macro_name = str(self.macro_list.get(selected[0]))
        if self._selected_macro_name in self.environment["macros"]:
            return self._selected_macro_name
        return None

    def selected_event_indices(self) -> list[int]:
        return sorted(int(item) for item in self.event_tree.selection())

    def selected_event_index(self) -> int | None:
        selected = self.selected_event_indices()
        return selected[0] if selected else None

    def select_event_indices(self, indices: list[int]) -> None:
        valid = [str(index) for index in indices if self.event_tree.exists(str(index))]
        if valid:
            self.event_tree.selection_set(valid)
            self.event_tree.focus(valid[0])
            self.event_tree.see(valid[0])

    def refresh_all(self) -> None:
        environments = list(self.data["environments"])
        self.environment_menu.configure(values=environments)
        self.environment_menu.set(self.data["current_env"])
        self.refresh_macros()
        self.reload_hotkeys()

    def refresh_macros(self) -> None:
        previous = self.selected_macro() or self._selected_macro_name
        query_text = self.search.get().strip().lower() if hasattr(self, "search") else ""
        self.macro_list.delete(0, "end")
        for name in sorted(self.environment["macros"], key=str.lower):
            if query_text in name.lower():
                self.macro_list.insert("end", name)
        names = self.macro_list.get(0, "end")
        if previous in names:
            index = names.index(previous)
            self.macro_list.selection_set(index)
            self.macro_list.activate(index)
            self.macro_list.see(index)
            self._selected_macro_name = previous

    def show_macro(self) -> None:
        name = self.selected_macro()
        if not name:
            return
        macro = self.environment["macros"][name]
        loops = int(macro.get("loops", 1))
        loop_label = "continuous" if loops == 0 else str(loops)
        self.macro_title.configure(text=f"{name}   |   Hotkey: {macro.get('hotkey', 'None')}   |   Loops: {loop_label}   |   Speed: {macro.get('speed', 1)}x")
        self.event_tree.delete(*self.event_tree.get_children())
        for index, event in enumerate(macro.get("events", [])):
            self.event_tree.insert("", "end", iid=str(index), values=(index + 1, f"{safe_float(event.get('dt', 0)):.3f}s", event_summary(event)))

    def new_macro(self) -> None:
        name = simpledialog.askstring("New macro", "Macro name:", parent=self)
        if not name:
            return
        name = name.strip()
        if not name or name in self.environment["macros"]:
            messagebox.showwarning(APP_NAME, "That macro name is empty or already exists.")
            return
        self.environment["macros"][name] = {"events": [], "hotkey": "None", "mouse": "None", "speed": 1.0, "loops": 1}
        self._selected_macro_name = name
        self.autosave()
        self.refresh_macros()
        names = self.macro_list.get(0, "end")
        if name in names:
            index = names.index(name)
            self.macro_list.selection_set(index)
            self.show_macro()

    def rename_macro(self) -> None:
        old_name = self.selected_macro()
        if not old_name:
            messagebox.showinfo(APP_NAME, "Select a macro first.")
            return
        new_name = simpledialog.askstring("Rename macro", "New macro name:", initialvalue=old_name, parent=self)
        if new_name is None:
            return
        new_name = new_name.strip()
        if not new_name or (new_name != old_name and new_name in self.environment["macros"]):
            messagebox.showwarning(APP_NAME, "That macro name is empty or already exists.")
            return
        if new_name == old_name:
            return
        macros = self.environment["macros"]
        macros[new_name] = macros.pop(old_name)
        for macro in macros.values():
            for event in macro.get("events", []):
                if event.get("type") == "play_macro" and event.get("macro_name") == old_name:
                    event["macro_name"] = new_name
        self._selected_macro_name = new_name
        self.refresh_macros()
        names = self.macro_list.get(0, "end")
        if new_name in names:
            index = names.index(new_name)
            self.macro_list.selection_clear(0, "end")
            self.macro_list.selection_set(index)
            self.macro_list.activate(index)
            self.macro_list.see(index)
        self.reload_hotkeys()
        self.show_macro()
        self.autosave()

    def delete_macro(self) -> None:
        name = self.selected_macro()
        if name and messagebox.askyesno("Delete macro", f"Delete {name!r}?", parent=self):
            del self.environment["macros"][name]
            self.refresh_macros()
            self.event_tree.delete(*self.event_tree.get_children())
            self._selected_macro_name = None
            self.reload_hotkeys()
            self.autosave()

    def copy_macro(self) -> None:
        name = self.selected_macro()
        if not name:
            return
        new_name = simpledialog.askstring("Copy macro", "New macro name:", initialvalue=f"{name} Copy", parent=self)
        if new_name and new_name not in self.environment["macros"]:
            self.environment["macros"][new_name] = copy.deepcopy(self.environment["macros"][name])
            self.environment["macros"][new_name]["hotkey"] = "None"
            self._selected_macro_name = new_name
            self.refresh_macros()
            self.autosave()

    def add_event(self) -> None:
        name = self.selected_macro()
        if not name:
            messagebox.showinfo(APP_NAME, "Select or create a macro first.")
            return
        choice = simpledialog.askstring("Add event", "Enter one of:\n" + "\n".join(EVENT_TEMPLATES), parent=self)
        if choice:
            matches = {item.lower(): item for item in EVENT_TEMPLATES}
            selected = matches.get(choice.strip().lower())
            if selected:
                self.add_specific(selected)
            else:
                messagebox.showerror("Unknown event", "Use one of the displayed event names.")

    def add_specific(self, template_name: str) -> None:
        name = self.selected_macro()
        if not name:
            messagebox.showinfo(APP_NAME, "Select or create a macro first.")
            return
        events = self.environment["macros"][name]["events"]
        selected_index = self.selected_event_index()
        insertion_index = len(events) if selected_index is None else selected_index + 1
        dialog = EventDialog(self, EVENT_TEMPLATES[template_name], f"Add {template_name}")
        self.wait_window(dialog)
        if dialog.result is not None:
            events.insert(insertion_index, dialog.result)
            self.show_macro()
            self.autosave()
            self.event_tree.selection_set(str(insertion_index))
            self.event_tree.see(str(insertion_index))

    def edit_event(self) -> None:
        name = self.selected_macro()
        index = self.selected_event_index()
        if name is None or index is None:
            messagebox.showinfo(APP_NAME, "Select an event to edit.")
            return
        events = self.environment["macros"][name]["events"]
        dialog = EventDialog(self, events[index])
        self.wait_window(dialog)
        if dialog.result is not None:
            events[index] = dialog.result
            self.show_macro()
            self.autosave()
            self.event_tree.selection_set(str(index))
            self.event_tree.see(str(index))

    def duplicate_event(self) -> None:
        name = self.selected_macro()
        index = self.selected_event_index()
        if name is not None and index is not None:
            events = self.environment["macros"][name]["events"]
            events.insert(index + 1, copy.deepcopy(events[index]))
            self.show_macro()
            self.autosave()
            self.event_tree.selection_set(str(index + 1))

    def delete_event(self) -> None:
        name = self.selected_macro()
        indices = self.selected_event_indices()
        if name is None or not indices:
            return
        if len(indices) > 1 and not messagebox.askyesno("Delete events", f"Delete the {len(indices)} selected events?", parent=self):
            return
        events = self.environment["macros"][name]["events"]
        for index in reversed(indices):
            del events[index]
        next_index = min(indices[0], len(events) - 1)
        self.show_macro(); self.autosave()
        if next_index >= 0:
            self.select_event_indices([next_index])

    def move_event(self, direction: int) -> None:
        name = self.selected_macro(); indices = self.selected_event_indices()
        if name is None or not indices or direction not in (-1, 1):
            return
        events = self.environment["macros"][name]["events"]; selected = set(indices)
        if direction < 0:
            for index in indices:
                if index > 0 and index - 1 not in selected:
                    events[index - 1], events[index] = events[index], events[index - 1]
            new_indices = [max(0, index - 1) for index in indices]
        else:
            for index in reversed(indices):
                if index < len(events) - 1 and index + 1 not in selected:
                    events[index], events[index + 1] = events[index + 1], events[index]
            new_indices = [min(len(events) - 1, index + 1) for index in indices]
        if new_indices != indices:
            self.show_macro(); self.autosave(); self.select_event_indices(new_indices)

    def start_recording(self) -> None:
        if self.playback.playing:
            messagebox.showwarning(APP_NAME, "Stop playback before recording.")
            return
        if not self.selected_macro():
            messagebox.showinfo(APP_NAME, "Select or create a macro first.")
            return
        self.recorder.start()

    def finish_recording(self) -> None:
        captured = self.recorder.stop()
        name = self.selected_macro()
        if not name or not captured:
            self.status.set("No captured events to apply.")
            return
        choice = messagebox.askyesnocancel(
            "Apply recording",
            f"Captured {len(captured)} events.\n\nYes: replace current events\nNo: append after current events\nCancel: discard",
            parent=self,
        )
        if choice is True:
            self.environment["macros"][name]["events"] = captured
            self.autosave()
        elif choice is False:
            self.environment["macros"][name]["events"].extend(captured)
            self.autosave()
        self.show_macro()

    def play(self) -> None:
        name = self.selected_macro()
        if not name:
            messagebox.showinfo(APP_NAME, "Select a macro first.")
            return
        if self.recorder.active:
            messagebox.showwarning(APP_NAME, "Finish recording before playback.")
            return
        macro = self.environment["macros"][name]
        payload = {
            "macroName": name,
            "macros": copy.deepcopy(self.environment["macros"]),
            "legacyVariables": copy.deepcopy(self.environment.get("variables", {})),
            "trustedScripts": bool(self.data.get("trustedScripts", False)),
        }
        try:
            self.playback.start(payload)
            self.status.set(f"Running {name}. Escape stops immediately.")
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def _global_key_press(self, key: Any) -> None:
        if key == Key.esc:
            self.ui_queue.put(("escape", None))

    def emergency_stop(self) -> None:
        if self.recorder.active:
            self.recorder.stop()
        self.playback.kill()
        self.status.set("Stopped by Escape")

    def edit_hotkey(self) -> None:
        name = self.selected_macro()
        if not name:
            messagebox.showinfo(APP_NAME, "Select a macro first.")
            return
        dialog = HotkeyCaptureDialog(self, self.environment["macros"][name].get("hotkey", "None"))
        self.wait_window(dialog)
        if dialog.result is not None:
            try:
                if dialog.result != "None":
                    keyboard.HotKey.parse(dialog.result)
                    conflicts = [other_name for other_name, macro in self.environment["macros"].items()
                                 if other_name != name and str(macro.get("hotkey", "None")) == dialog.result]
                    if conflicts:
                        conflict_text = ", ".join(repr(item) for item in conflicts)
                        if not messagebox.askyesno("Hotkey already assigned", f"{dialog.result} is already assigned to {conflict_text}.\n\nAssign it here and clear it from the old macro?", parent=self):
                            return
                        for other_name in conflicts:
                            self.environment["macros"][other_name]["hotkey"] = "None"
                self.environment["macros"][name]["hotkey"] = dialog.result
                self.reload_hotkeys()
                self.show_macro()
                self.save()
            except Exception as exc:
                messagebox.showerror("Invalid hotkey", str(exc))

    def reset_hotkey(self) -> None:
        name = self.selected_macro()
        if not name:
            messagebox.showinfo(APP_NAME, "Select a macro first.")
            return
        self.environment["macros"][name]["hotkey"] = "None"
        self.reload_hotkeys()
        self.show_macro()
        self.autosave()

    def reload_hotkeys(self) -> None:
        if self.global_hotkeys:
            try:
                self.global_hotkeys.stop()
            except Exception:
                pass
            self.global_hotkeys = None
        mapping = {}
        duplicates = set()
        for name, macro in self.environment["macros"].items():
            hotkey = str(macro.get("hotkey", "None"))
            if hotkey and hotkey != "None":
                if hotkey in mapping:
                    duplicates.add(hotkey)
                mapping[hotkey] = lambda macro_name=name: self.ui_queue.put(("hotkey", macro_name))
        if duplicates:
            self.status.set("Duplicate hotkeys ignored: " + ", ".join(sorted(duplicates)))
            for duplicate in duplicates:
                mapping.pop(duplicate, None)
        if mapping:
            try:
                self.global_hotkeys = keyboard.GlobalHotKeys(mapping)
                self.global_hotkeys.start()
            except Exception as exc:
                self.status.set(f"Hotkey registration failed: {exc}")

    def edit_playback_settings(self) -> None:
        name = self.selected_macro()
        if not name:
            messagebox.showinfo(APP_NAME, "Select a macro first.")
            return
        macro = self.environment["macros"][name]
        loops = simpledialog.askinteger(
            "Loop count", "Number of loops. Enter 0 for continuous looping until Escape:",
            initialvalue=int(macro.get("loops", 1)), minvalue=0, parent=self,
        )
        if loops is None:
            return
        speed = simpledialog.askfloat(
            "Playback speed", "Speed multiplier, for example 1.0 or 2.0:",
            initialvalue=safe_float(macro.get("speed", 1), 1), minvalue=0.05, maxvalue=20, parent=self,
        )
        if speed is None:
            return
        macro["loops"] = loops
        macro["speed"] = speed
        self.show_macro()
        self.autosave()

    def add_environment(self) -> None:
        name = simpledialog.askstring("New environment", "Environment name:", parent=self)
        if name and name.strip() and name.strip() not in self.data["environments"]:
            self.data["environments"][name.strip()] = {"macros": {}, "variables": {}}
            self.data["current_env"] = name.strip()
            self._selected_macro_name = None
            self.refresh_all()
            self.autosave()

    def change_environment(self, name: str) -> None:
        if self.playback.playing:
            self.environment_menu.set(self.data["current_env"])
            messagebox.showwarning(APP_NAME, "Press Escape before changing environments.")
            return
        self.data["current_env"] = name
        self._selected_macro_name = None
        self.refresh_macros()
        self.reload_hotkeys()
        self.event_tree.delete(*self.event_tree.get_children())
        self.macro_title.configure(text="Select a macro")
        self.autosave()

    def autosave(self) -> None:
        if self._autosave_job is not None:
            try: self.after_cancel(self._autosave_job)
            except Exception: pass
        self._autosave_job = self.after(250, self.save)

    def save(self) -> None:
        try:
            self.storage.save(self.data)
            self.status.set(f"Auto-saved. Backup: {SAVE_FILE.name}.bak")
        except Exception as exc:
            logger.exception("Save failed")
            messagebox.showerror("Save failed", f"Your existing save file was not replaced.\n\n{exc}", parent=self)

    def report(self, text: str) -> None:
        self.ui_queue.put(("status", text))

    def drain_queue(self) -> None:
        try:
            while True:
                kind, value = self.ui_queue.get_nowait()
                if kind == "escape":
                    self.emergency_stop()
                elif kind == "hotkey":
                    if not self.playback.playing:
                        self._play_by_name(str(value))
                elif kind == "status":
                    self.status.set(str(value))
                elif kind == "done":
                    self.status.set(str(value))
                    if str(value).startswith("Failed"):
                        messagebox.showerror("Playback failed", str(value), parent=self)
        except queue.Empty:
            pass
        self.after(75, self.drain_queue)

    def _play_by_name(self, name: str) -> None:
        names = self.macro_list.get(0, "end")
        if name in names:
            index = names.index(name)
            self.macro_list.selection_clear(0, "end")
            self.macro_list.selection_set(index)
            self.macro_list.see(index)
            self.show_macro()
            self.play()

    def close(self) -> None:
        self.playback.kill()
        if self.recorder.active:
            self.recorder.stop()
        if self.global_hotkeys:
            try:
                self.global_hotkeys.stop()
            except Exception:
                pass
        try:
            self.escape_listener.stop()
        except Exception:
            pass
        try:
            self.storage.save(self.data)
        except Exception:
            logger.exception("Save on exit failed")
        self.destroy()


if __name__ == "__main__":
    mp.freeze_support()
    VisionSuiteApp().mainloop()
