import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


def _base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


def _setup() -> logging.Logger:
    log_dir = _base_dir() / 'logs'
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / 'automation.log'

    handler = RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=3,
        encoding='utf-8',
    )
    handler.setFormatter(logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    ))

    log = logging.getLogger('gelatik')
    log.setLevel(logging.DEBUG)
    if not log.handlers:
        log.addHandler(handler)
    return log


logger = _setup()
