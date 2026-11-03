"""Test setup: stub Windows-only COM modules and keep config/log writes in temp dirs."""
import os
import sys
import types
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

for name in ("win32com", "win32com.client", "pythoncom"):
    if name not in sys.modules:
        try:
            __import__(name)
        except ImportError:
            mod = types.ModuleType(name)
            mod.CoInitialize = mod.CoUninitialize = lambda *a, **k: None
            mod.Dispatch = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("COM unavailable in tests"))
            sys.modules[name] = mod
sys.modules["win32com"].client = sys.modules["win32com.client"]


@pytest.fixture(autouse=True)
def isolated_config(tmp_path, monkeypatch):
    from utils import app_utils
    monkeypatch.setattr(app_utils, "CONFIG_PATH", str(tmp_path / "rfp_config.json"))
    yield
