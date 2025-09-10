"""Helpers for backing up existing export files.

Centralizes move+timestamp logic used by overview export and the workflow-based exports.
"""
from __future__ import annotations

from datetime import datetime
from enum import Enum
import json
import shutil
from pathlib import Path

from loguru import logger

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import Directory


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def timestamped_filename(original: Path, timestamp: datetime | None = None) -> str:
    ts = (timestamp or datetime.utcnow()).strftime("%Y%m%d_%H%M%S")
    return f"{original.stem}_{ts}{original.suffix}"


def backup_existing_file(target_path: Path, backups_dir: Path, manifest: dict | None = None) -> Path:
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


class BackupFlag(Enum):
    BACKUP = "backup"
    RESTORE = "restore"
    NONE = "none"
    DEFAULT = "default"


class RestoreOptions(Enum):
    LATEST = "latest"
    OLDEST = "oldest"
    MANUAL = "manual"


class RestoreStrategy(Enum):
    REPLACE = "replace"
    MERGE_PREFER_EXISTING = "merge_prefer_existing"
    MERGE_PREFER_BACKUP = "merge_prefer_backup"


class Backupper:
    def __init__(self) -> None:
        """Utilities for creating and restoring backups.

        The CLI hooks into this class; backup configuration lives in settings.yaml
        under the `backup` and `directories` sections.
        """
        ...

    def backup_files(self) -> None:
        dirs_to_backup = SETTINGS.backup_settings.backup_dirs
        backup_location = SETTINGS.backup_settings.backup_location
        max_backups = SETTINGS.backup_settings.max_backups

        if not dirs_to_backup:
            logger.warning(
                "backup_all is set to true in settings.yaml, but no dirs to backup were specified. Skipping."
            )
            return
        if not backup_location:
            logger.warning(
                "backup_all is set to true in settings.yaml, but no backup location was specified. Skipping."
            )
            return
        if not max_backups:
            logger.warning(
                "backup_all is set to true in settings.yaml, but no max amount of backups was specified. Skipping."
            )
            return
        backup_location_dirs = backup_location.dirs()
        if len(backup_location_dirs) >= max_backups:
            logger.info(
                f"Found {len(backup_location_dirs)} backups, making room by deleting oldest backup(s)."
            )
            while len(backup_location_dirs) >= max_backups:
                dir_by_date = {d.created: d for d in backup_location_dirs if d.created}
                dates = list(dir_by_date.keys())
                dates.sort()
                dir_by_date[dates[0]].delete()
                backup_location_dirs = backup_location.dirs()

        logger.info(
            f"Creating backup of all data in dirs: {[d.full.name for d in dirs_to_backup]}"
        )
        for i in range(0, max_backups + 4):
            new_backup_dir = Directory(
                backup_location.full
                / f"backup{i if i > 0 else ''}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}",
                create_dir=False,
            )
            if not new_backup_dir.exists:
                new_backup_dir.create()
                break

        for d in dirs_to_backup:
            d.copy(new_backup_dir)

        logger.success(f"Backups done, stored in {new_backup_dir.full}")

    def restore_backup(
        self,
        backup_dir: Directory | None = SETTINGS.backup_settings.backup_location,
        strategy: RestoreStrategy = RestoreStrategy.REPLACE,
        select: RestoreOptions = RestoreOptions.LATEST,
    ) -> None:
        """Restore a backup from a directory. Defaults to restoring the latest backup.

        Parameters:
            backup_dir: Directory: the directory to restore from. If not specified, the dir in backup_settings will be used.
            strategy: RestoreStrategy: the strategy to use for restoring the backup.
            select: RestoreOptions: Which backup to restore.
        """

        target_dir = SETTINGS.dirs[DirSetting.FACULTIES_DIR]

        if not backup_dir:
            backup_dir = SETTINGS.backup_settings.backup_location
        if backup_dir is None:
            logger.warning("No backup location configured; cannot restore backup.")
            return
        if not isinstance(backup_dir, Directory) and backup_dir:
            try:
                backup_dir = Directory(backup_dir)
            except Exception:
                logger.warning(
                    f"Could not convert {backup_dir} to a Directory object. Cannot restore backup."
                )
                return
        if not backup_dir.exists:
            logger.warning(
                f"Backup dir {backup_dir.full} does not exist -- cannot restore backup."
            )
            return

        selected_backup_dir = None
        if select == RestoreOptions.LATEST:
            selected_backup_dir = max(backup_dir.dirs(), key=lambda x: x.created)
        elif select == RestoreOptions.OLDEST:
            selected_backup_dir = min(backup_dir.dirs(), key=lambda x: x.created)
        elif select == RestoreOptions.MANUAL:
            logger.info(f"  Select backup to restore from {backup_dir.full}:")
            logger.info("------------------------------------------------------\n")
            for num, dir in enumerate(backup_dir.dirs()):
                logger.info(f"    {num}: {dir.full}")
            logger.info("\n")
            while not selected_backup_dir:
                try:
                    selected_backup_dir = backup_dir.dirs()[
                        int(
                            input(
                                f"  Select backup to restore (0-{len(backup_dir.dirs()) - 1}): "
                            )
                        )
                    ]
                except Exception as e:
                    logger.warning(
                        f"Invalid input. Please select a valid backup. ({e})"
                    )

        if not selected_backup_dir:
            logger.warning("No valid backup selected, cannot restore.")
            return

        if strategy == RestoreStrategy.REPLACE:
            target_dir.delete()
            logger.info(f"Deleted {target_dir.full}")
            logger.info(
                f"Restoring backup from {selected_backup_dir.full} to {target_dir.full} using strategy: {strategy}"
            )
            selected_backup_dir.copy(target_dir)
        elif strategy == RestoreStrategy.MERGE_PREFER_BACKUP:
            selected_backup_dir.copy(target_dir)
        elif strategy == RestoreStrategy.MERGE_PREFER_EXISTING:
            selected_backup_dir.copy(target_dir, overwrite=False)
        else:
            logger.warning(f"Unrecognized strategy: {strategy}. Cannot restore backup.")
            return
        logger.success(
            f"Backup restored to {target_dir.full} using strategy: {strategy}"
        )
