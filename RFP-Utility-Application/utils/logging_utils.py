import logging
import os
import sys
import time
import uuid
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler
from typing import Any, Dict, Iterator, Optional


_LOGGER_READY = False
_LOG_FILE_PATH = None
_QMESSAGEBOX_HOOKED = False
_DEFAULT_FIELDS: Dict[str, str] = {"run_id": "-"}


class _DefaultFieldsFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        # Ensure format fields exist even if callsites don't pass extra=...
        for k, v in _DEFAULT_FIELDS.items():
            if not hasattr(record, k):
                setattr(record, k, v)
        return True


def _app_directory() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # Mirror the app's existing config/data location in source runs (utils folder).
    return os.path.dirname(os.path.abspath(__file__))


def get_log_file_path() -> str:
    global _LOG_FILE_PATH
    if _LOG_FILE_PATH:
        return _LOG_FILE_PATH
    _LOG_FILE_PATH = os.path.join(_app_directory(), "myagencypro.log")
    return _LOG_FILE_PATH


def configure_app_logging() -> str:
    global _LOGGER_READY
    if _LOGGER_READY:
        return get_log_file_path()

    log_path = get_log_file_path()
    os.makedirs(os.path.dirname(log_path), exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | run=%(run_id)s | %(filename)s:%(lineno)d | %(message)s"
    )

    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    file_handler.addFilter(_DefaultFieldsFilter())

    root.handlers.clear()
    root.addHandler(file_handler)

    logging.captureWarnings(True)

    def _handle_uncaught(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logging.getLogger("uncaught").exception(
            "Uncaught exception",
            exc_info=(exc_type, exc_value, exc_traceback),
        )

    sys.excepthook = _handle_uncaught

    _LOGGER_READY = True
    logging.getLogger(__name__).info("Logging initialized at %s", log_path)
    return log_path


def get_logger(name: str) -> logging.Logger:
    if not _LOGGER_READY:
        configure_app_logging()
    return logging.getLogger(name)


def new_run_id() -> str:
    return uuid.uuid4().hex[:10]


@contextmanager
def log_timing(logger: logging.Logger, msg: str, *, run_id: Optional[str] = None, level: int = logging.INFO) -> Iterator[None]:
    start = time.perf_counter()
    try:
        yield
    finally:
        elapsed_ms = int((time.perf_counter() - start) * 1000)
        extra = {"run_id": run_id} if run_id else None
        logger.log(level, "%s | elapsed_ms=%s", msg, elapsed_ms, extra=extra)


def _safe_text(value: Any) -> str:
    try:
        return str(value)
    except Exception:
        return "<unprintable>"


def install_qmessagebox_logging() -> None:
    global _QMESSAGEBOX_HOOKED
    if _QMESSAGEBOX_HOOKED:
        return

    try:
        from PySide6.QtWidgets import QMessageBox
    except Exception:
        return

    logger = get_logger("ui.messagebox")

    def _wrap(method_name: str, level: int):
        original = getattr(QMessageBox, method_name)

        def wrapper(*args, **kwargs):
            title = _safe_text(args[1]) if len(args) > 1 else _safe_text(kwargs.get("title", ""))
            text = _safe_text(args[2]) if len(args) > 2 else _safe_text(kwargs.get("text", ""))
            logger.log(level, "%s | %s | %s", method_name.upper(), title, text)
            return original(*args, **kwargs)

        return staticmethod(wrapper)

    QMessageBox.warning = _wrap("warning", logging.WARNING)
    QMessageBox.critical = _wrap("critical", logging.ERROR)
    QMessageBox.information = _wrap("information", logging.INFO)

    _QMESSAGEBOX_HOOKED = True
    logger.info("QMessageBox logging hook installed")
