"""Loads and validates config.yaml."""

import logging
from pathlib import Path

import yaml

logger = logging.getLogger(__name__)

REQUIRED_KEYS = [
    "input_folder",
    "output_folder",
    "template_path",
    "required_files",
    "photo_settings",
    "chart_settings",
    "report_settings",
    "file_patterns",
    "file_extensions",
]


def load_config(config_path: str | Path = "config.yaml") -> dict:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path.resolve()}")

    with open(path, "r") as f:
        cfg = yaml.safe_load(f)

    if cfg is None:
        raise ValueError(f"Config file is empty: {path.resolve()}")

    missing = [k for k in REQUIRED_KEYS if k not in cfg]
    if missing:
        raise ValueError(f"Config is missing required keys: {missing}")

    logger.debug("Config loaded from %s", path.resolve())
    return cfg
