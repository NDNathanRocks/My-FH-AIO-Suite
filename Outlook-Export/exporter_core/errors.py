"""Failure classification: every failure gets a stage, category and retry flag."""
from __future__ import annotations

import traceback
from dataclasses import dataclass

STAGES = ("snapshot", "open_item", "read_properties", "build_path", "save_temp", "move", "verify", "db_write", "adopt")


@dataclass(frozen=True)
class FailureInfo:
    stage: str
    category: str
    error_type: str
    message: str
    hresult: str
    retryable: bool

    def short(self) -> str:
        return f"[{self.stage}/{self.category}] {self.error_type}: {self.message}"[:500]


class StageError(Exception):
    """Wraps an exception with the processing stage where it happened."""

    def __init__(self, stage: str, original: BaseException):
        super().__init__(str(original))
        self.stage, self.original = stage, original


def exception_details(exc: BaseException) -> dict:
    details = {"type": type(exc).__name__, "message": str(exc), "repr": repr(exc),
               "traceback": "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))}
    for name in ("hresult", "strerror", "excepinfo", "argerror"):
        if hasattr(exc, name):
            details[name] = repr(getattr(exc, name))
    return details


def is_memory_error(exc: BaseException) -> bool:
    text = (str(exc) + repr(exc)).lower()
    return any(value in text for value in ("not enough memory", "-2147024882", "0x8007000e", "4096"))


_RULES = (
    # (category, retryable, substrings)
    ("memory", True, ("not enough memory", "-2147024882", "0x8007000e", "memoryerror")),
    ("disk_full", True, ("no space left", "disk full", "insufficient output free space", "-2147024784", "errno 28")),
    ("path_too_long", False, ("path too long", "filename too long", "winerror 206", "-2147024690")),
    ("permission", True, ("access is denied", "permission denied", "winerror 5", "-2147024891", "errno 13")),
    ("file_locked", True, ("being used by another process", "winerror 32", "sharing violation", "-2147024864")),
    ("item_not_found", False, ("could not be found", "not found", "-2147221233", "mapi_e_not_found")),
    ("corrupt_item", False, ("corrupt", "-2147221221", "mapi_e_corrupt_data", "cannot be opened")),
    ("empty_output", True, ("missing or empty",)),
    ("outlook_busy", True, ("call was rejected", "rpc_e_call_rejected", "-2147418111", "server busy", "-2147417846", "rpc server is unavailable", "-2147023174")),
)


def classify(exc: BaseException, stage: str | None = None) -> FailureInfo:
    original = exc.original if isinstance(exc, StageError) else exc
    stage = stage or (exc.stage if isinstance(exc, StageError) else "unknown")
    text = (str(original) + " " + repr(original) + " " + repr(getattr(original, "hresult", ""))).lower()
    category, retryable = "unknown", True
    for name, can_retry, needles in _RULES:
        if any(n in text for n in needles):
            category, retryable = name, can_retry
            break
    else:
        if isinstance(original, MemoryError):
            category = "memory"
        elif isinstance(original, PermissionError):
            category = "permission"
        elif isinstance(original, FileNotFoundError):
            category = "path_missing"
        elif type(original).__name__ == "com_error":
            category = "com_error"
    return FailureInfo(stage, category, type(original).__name__, str(original)[:1000],
                       repr(getattr(original, "hresult", "")) if hasattr(original, "hresult") else "", retryable)
