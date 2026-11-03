"""Single-instance run lock (prevents two exporters writing the same database)."""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path


def pid_running(pid: int) -> bool:
    if not pid:
        return False
    if sys.platform.startswith("win"):
        import ctypes
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    try:
        os.kill(int(pid), 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def acquire_lock(lock_path: Path) -> None:
    if lock_path.exists():
        try:
            info = json.loads(lock_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            info = {"pid": 0}
        pid = int(info.get("pid") or 0)
        if pid_running(pid) and pid != os.getpid():
            raise RuntimeError(f"Another exporter is active (PID {pid}). Close it or wait for it to finish.")
        lock_path.unlink(missing_ok=True)  # stale lock from a crashed run
    fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    try:
        os.write(fd, json.dumps({"pid": os.getpid(), "created": datetime.now().isoformat()}).encode())
    finally:
        os.close(fd)


def release_lock(lock_path: Path) -> None:
    try:
        info = json.loads(lock_path.read_text(encoding="utf-8"))
        if int(info.get("pid") or 0) not in (0, os.getpid()):
            return  # not ours
    except (OSError, ValueError):
        pass
    lock_path.unlink(missing_ok=True)
