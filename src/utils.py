"""Shared utility helpers."""

import logging
import sys
from pathlib import Path


def setup_logging(level: str = "INFO") -> None:
    numeric = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=numeric,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def find_files_by_prefix(folder: Path, prefixes: list[str], extensions: list[str]) -> list[Path]:
    """Return files whose names start with any prefix and end with any extension (case-insensitive)."""
    matches = []
    prefixes_lower = [p.lower() for p in prefixes]
    exts_lower = [e.lower() for e in extensions]
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        name_lower = f.name.lower()
        if any(name_lower.startswith(p) for p in prefixes_lower) and f.suffix.lower() in exts_lower:
            matches.append(f)
    return matches


def sanitize_for_filename(value: str) -> str:
    """Replace spaces and special characters for safe use in file names."""
    safe = "".join(c if c.isalnum() or c in "-_." else "" for c in value.replace(" ", ""))
    return safe
