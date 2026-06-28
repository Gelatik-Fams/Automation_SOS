import json
import sys
from pathlib import Path


def _base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


BASE_DIR = _base_dir()
_CONFIG_PATH = BASE_DIR / 'config.json'
_VERSION_PATH = BASE_DIR / 'VERSION'

_DEFAULTS = {
    'theme': 'light',
    'last_folder': '',
    'window_width': 860,
    'window_height': 620,
    'version': '1.0.0',
}


def get_version() -> str:
    if _VERSION_PATH.exists():
        v = _VERSION_PATH.read_text(encoding='utf-8').strip()
        if v:
            return v
    return _DEFAULTS['version']


def load_config() -> dict:
    cfg = _DEFAULTS.copy()
    cfg['version'] = get_version()
    if _CONFIG_PATH.exists():
        try:
            data = json.loads(_CONFIG_PATH.read_text(encoding='utf-8'))
            cfg.update({k: v for k, v in data.items() if k in _DEFAULTS})
        except Exception:
            pass
    return cfg


def save_config(cfg: dict) -> None:
    try:
        _CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding='utf-8')
    except Exception:
        pass
