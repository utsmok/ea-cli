"""Helpers for backing up existing export files.

Centralizes move+timestamp logic used by overview export and the workflow-based exports.
"""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path

from loguru import logger


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def timestamped_filename(original: Path, timestamp: datetime | None = None) -> str:
    ts = (timestamp or datetime.utcnow()).strftime("%Y%m%d_%H%M%S")
    return f"{original.stem}_backup_{ts}{original.suffix}"


def backup_existing_file(
    target_path: Path, backups_dir: Path, manifest: dict | None = None
) -> Path:
    """Move ``target_path`` into ``backups_dir`` and return the moved path.

    If the target doesn't exist, the original Path is returned unchanged.
    """
    if not target_path.exists():
        return target_path

    ensure_dir(backups_dir)
    new_name = timestamped_filename(target_path)
    dest = backups_dir / new_name
    # Use shutil.move to preserve perms where possible
    shutil.move(str(target_path), str(dest))

    # Write optional manifest next to the moved file (best-effort)
    if manifest is not None:
        manifest_path = dest.with_suffix(dest.suffix + ".manifest.json")
        try:
            with manifest_path.open("w", encoding="utf-8") as fh:
                json.dump(manifest, fh, ensure_ascii=False, indent=2)
        except Exception:
            logger.debug("Failed to write backup manifest; continuing without manifest")

    return dest
