"""
This module provides functionality for backing up and restoring application data,
primarily focusing on directories specified in the settings (e.g., faculty sheets).
It supports strategies like full replacement or merging, and selection of backups
by recency or manual choice.
"""

import logging
import traceback
from datetime import datetime
from enum import Enum

from easy_access.settings import SETTINGS, DirSetting
from easy_access.utils import Directory  # Removed cool, info, print, warn

logger = logging.getLogger(__name__)


class BackupFlag(Enum):
    """Specifies the backup/restore action to be taken when the tool starts."""

    BACKUP = "backup"  # Perform a backup.
    RESTORE = "restore"  # Perform a restore.
    NONE = "none"  # Do nothing related to backup/restore.
    DEFAULT = "default"  # Use behavior defined in settings.yaml (e.g., backup_all).


class RestoreOptions(Enum):
    """Defines which backup to select when restoring."""

    LATEST = "latest"  # Restore the most recent backup.
    OLDEST = "oldest"  # Restore the oldest available backup.
    MANUAL = "manual"  # Allow manual selection from a list of available backups.


class RestoreStrategy(Enum):
    """Defines the strategy to use when restoring a backup."""

    REPLACE = "replace"  # Completely replace the target directory with the backup.
    MERGE_PREFER_EXISTING = "merge_prefer_existing"  # Merge: keep existing files if conflicts, add new from backup.
    MERGE_PREFER_BACKUP = "merge_prefer_backup"  # Merge: overwrite existing files with backup's if conflicts, add new.


class Backupper:
    """
    Handles backup and restore operations for specified application directories.
    Configuration for backups (paths, strategy, limits) is drawn from the
    global SETTINGS object, typically loaded from `settings.yaml`.
    """

    def __init__(self) -> None:
        """
        Initializes the Backupper.
        Backup settings are read from the global SETTINGS object.
        """
        # Docstring from original code moved to class level for clarity.
        # The __init__ itself doesn't do much beyond instantiation.
        pass

    def backup_files(self) -> None:
        """
        Creates a backup of specified directories.

        It checks `SETTINGS.backup_settings` for directories to back up, the backup
        location, and the maximum number of backups to retain. If the number of
        existing backups exceeds the maximum, the oldest ones are deleted.
        A new backup directory is created with a timestamp.
        """
        dirs_to_backup: list[Directory] = SETTINGS.backup_settings.backup_dirs
        backup_location: Directory | None = SETTINGS.backup_settings.backup_location
        max_backups: int = SETTINGS.backup_settings.max_backups

        if not dirs_to_backup:
            logger.warning(
                "backup_all is True (or backup forced) but no backup_dirs specified in settings. Skipping backup."
            )
            return
        if not backup_location:
            logger.warning(
                "backup_all is True (or backup forced) but no backup_location specified in settings. Skipping backup."
            )
            return
        if not backup_location.exists:  # Ensure base backup location exists
            try:
                backup_location.create()
                logger.info(
                    f"Created main backup location directory: {backup_location.full}"
                )
            except Exception as e:
                logger.error(
                    f"Failed to create main backup location {backup_location.full}: {e}. Skipping backup."
                )
                return

        if max_backups <= 0:  # max_backups should be positive
            logger.warning(
                f"max_backups is {max_backups}, which is invalid or disables backup retention. Skipping backup."
            )
            return

        # Manage existing backups: delete oldest if count exceeds max_backups
        try:
            existing_backup_dirs = [
                d for d in backup_location.dirs if d.name.startswith("backup_")
            ]  # Filter for actual backup dirs
            while len(existing_backup_dirs) >= max_backups:
                if not existing_backup_dirs:
                    break  # Should not happen if len >= max_backups > 0
                oldest_backup = min(existing_backup_dirs, key=lambda d: d.created)
                logger.info(
                    f"Max backups ({max_backups}) reached. Deleting oldest backup: {oldest_backup.full}"
                )
                oldest_backup.delete()
                existing_backup_dirs = [
                    d for d in backup_location.dirs if d.name.startswith("backup_")
                ]
        except Exception as e:
            logger.error(
                f"Error managing existing backups: {e}. Proceeding with new backup if possible."
            )

        # Create new backup directory with a unique timestamped name
        timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        new_backup_dir_name = f"backup_{timestamp_str}"
        new_backup_dir = Directory(
            backup_location.full / new_backup_dir_name, create_dir=False
        )

        # Fallback for unique name if somehow multiple backups are made in the same second (highly unlikely)
        counter = 0
        while new_backup_dir.exists:  # pragma: no cover
            counter += 1
            new_backup_dir_name = f"backup_{timestamp_str}_{counter}"
            new_backup_dir = Directory(
                backup_location.full / new_backup_dir_name, create_dir=False
            )

        try:
            new_backup_dir.create()
            logger.info(f"Created new backup directory: {new_backup_dir.full}")
        except Exception as e:
            logger.error(
                f"Failed to create new backup directory {new_backup_dir.full}: {e}. Skipping backup."
            )
            return

        # Copy specified directories to the new backup location
        dir_names_to_backup_str = ", ".join([d.name for d in dirs_to_backup if d])
        logger.info(
            f"Starting backup of directories: [{dir_names_to_backup_str}] to {new_backup_dir.name}"
        )
        for dir_to_backup in dirs_to_backup:
            if dir_to_backup and dir_to_backup.exists:
                try:
                    # The target for copytree should be a subdirectory within new_backup_dir, named after the source dir
                    target_copy_path = new_backup_dir.full / dir_to_backup.name
                    dir_to_backup.copy(
                        target_copy_path, overwrite=True
                    )  # Directory.copy handles its own logging
                except Exception as e_copy:
                    logger.error(
                        f"Failed to copy directory {dir_to_backup.name} to backup: {e_copy}"
                    )
            else:
                logger.warning(
                    f"Directory {str(dir_to_backup.name if dir_to_backup else 'N/A')} not found or invalid. Skipping its backup."
                )

        logger.info(f"Backup process completed. Data saved in {new_backup_dir.full}")

    def restore_backup(
        self,
        backup_source_location: Directory
        | None = None,  # Renamed from backup_dir for clarity
        strategy: RestoreStrategy = RestoreStrategy.REPLACE,
        select: RestoreOptions = RestoreOptions.LATEST,
    ) -> None:
        """
        Restores data from a backup to the configured 'faculties_dir'.

        Args:
            backup_source_location (Directory, optional): The base directory containing multiple timestamped backups.
                                                        If None, uses `SETTINGS.backup_settings.backup_location`.
            strategy (RestoreStrategy, optional): The strategy for restoring ('replace', 'merge_prefer_backup',
                                                'merge_prefer_existing'). Defaults to RestoreStrategy.REPLACE.
            select (RestoreOptions, optional): Which specific backup to restore ('latest', 'oldest', 'manual').
                                             Defaults to RestoreOptions.LATEST.
        """
        target_restore_dir: Directory | None = SETTINGS.dirs.get(
            DirSetting.FACULTIES_DIR
        )
        if not target_restore_dir:
            logger.error(
                "Target directory for restore ('faculties_dir') not configured in settings. Cannot restore."
            )
            return

        actual_backup_source_location: Directory | None = (
            backup_source_location or SETTINGS.backup_settings.backup_location
        )
        if not actual_backup_source_location or not isinstance(
            actual_backup_source_location, Directory
        ):
            logger.error(
                f"Invalid backup source location provided or configured: {actual_backup_source_location}"
            )
            return
        if not actual_backup_source_location.exists:
            logger.error(
                f"Backup source location {actual_backup_source_location.full} does not exist. Cannot restore."
            )
            return

        available_backups: list[Directory] = [
            d
            for d in actual_backup_source_location.dirs
            if d.name.startswith("backup_")
        ]
        if not available_backups:
            logger.warning(
                f"No backups found in {actual_backup_source_location.full}. Cannot restore."
            )
            return

        selected_backup_to_restore_from: Directory | None = None
        if select == RestoreOptions.LATEST:
            selected_backup_to_restore_from = max(
                available_backups, key=lambda d: d.created
            )
        elif select == RestoreOptions.OLDEST:
            selected_backup_to_restore_from = min(
                available_backups, key=lambda d: d.created
            )
        elif select == RestoreOptions.MANUAL:  # pragma: no cover
            logger.info(f"Available backups in {actual_backup_source_location.full}:")
            for i, bk_dir in enumerate(available_backups):
                logger.info(
                    f"  {i}: {bk_dir.name} (Created: {bk_dir.created.strftime('%Y-%m-%d %H:%M:%S')})"
                )

            while True:
                try:
                    choice_str = input(
                        f"  Select backup to restore by number (0-{len(available_backups) - 1}): "
                    )
                    choice_idx = int(choice_str)
                    if 0 <= choice_idx < len(available_backups):
                        selected_backup_to_restore_from = available_backups[choice_idx]
                        break
                    else:
                        logger.warning(
                            "Invalid selection. Please enter a valid number."
                        )
                except ValueError:
                    logger.warning("Invalid input. Please enter a number.")
                except Exception as e_input:  # Catch other potential input errors
                    logger.error(f"Error during manual selection: {e_input}")
                    return  # Exit if selection fails unexpectedly

        if (
            not selected_backup_to_restore_from
            or not selected_backup_to_restore_from.exists
        ):
            logger.error(
                f"Selected backup ({selected_backup_to_restore_from.name if selected_backup_to_restore_from else 'None'}) is not valid or does not exist. Cannot restore."
            )
            return

        logger.info(
            f"Selected backup for restore: {selected_backup_to_restore_from.full}"
        )
        logger.info(f"Target directory for restore: {target_restore_dir.full}")
        logger.info(f"Restore strategy: {strategy.value}")

        try:
            if strategy == RestoreStrategy.REPLACE:
                if target_restore_dir.exists:
                    logger.info(
                        f"Strategy '{strategy.value}': Deleting target directory {target_restore_dir.full} before restore."
                    )
                    target_restore_dir.delete()
                # Copy contents of each item within selected_backup_to_restore_from into target_restore_dir
                # Example: if backup is backup_xxx/faculties_dir_content/, copy faculties_dir_content/* to target_dir
                for item_in_backup in (
                    selected_backup_to_restore_from.dirs
                ):  # Assuming backup stores backed up dirs as subdirs
                    target_path_for_item = target_restore_dir.full / item_in_backup.name
                    logger.info(
                        f"Restoring {item_in_backup.name} to {target_path_for_item}..."
                    )
                    item_in_backup.copy(target_path_for_item, overwrite=True)
                # If the backup dir itself IS the content (e.g. backup_xxx is a copy of faculties_dir)
                # then the copy call would be: selected_backup_to_restore_from.copy(target_restore_dir.full, overwrite=True)
                # Assuming the former structure based on `d.copy(new_backup_dir)` in backup_files.
                # This means new_backup_dir contains subdirs like 'faculties_dir', 'other_backed_up_dir'.
                # So when restoring, we need to iterate through these and copy them to their original locations.
                # The current target_dir is SETTINGS.dirs[DirSetting.FACULTIES_DIR].
                # So, if selected_backup_to_restore_from contains a 'faculties_dir', that's what we copy.
                backed_up_faculties_content = Directory(
                    selected_backup_to_restore_from.full / target_restore_dir.name,
                    create_dir=False,
                )
                if backed_up_faculties_content.exists:
                    logger.info(
                        f"Restoring content from {backed_up_faculties_content.full} to {target_restore_dir.full}"
                    )
                    backed_up_faculties_content.copy(
                        target_restore_dir.full, overwrite=True
                    )
                else:
                    logger.warning(
                        f"Expected content for '{target_restore_dir.name}' not found in backup {selected_backup_to_restore_from.name}. Target may be empty."
                    )

            elif strategy == RestoreStrategy.MERGE_PREFER_BACKUP:
                backed_up_faculties_content = Directory(
                    selected_backup_to_restore_from.full / target_restore_dir.name,
                    create_dir=False,
                )
                if backed_up_faculties_content.exists:
                    logger.info(
                        f"Merging (preferring backup) from {backed_up_faculties_content.full} into {target_restore_dir.full}"
                    )
                    backed_up_faculties_content.copy(
                        target_restore_dir.full, overwrite=True
                    )  # overwrite=True prefers source (backup)
                else:
                    logger.warning(
                        f"Expected content for '{target_restore_dir.name}' not found in backup. Nothing to merge."
                    )

            elif strategy == RestoreStrategy.MERGE_PREFER_EXISTING:
                backed_up_faculties_content = Directory(
                    selected_backup_to_restore_from.full / target_restore_dir.name,
                    create_dir=False,
                )
                if backed_up_faculties_content.exists:
                    logger.info(
                        f"Merging (preferring existing) from {backed_up_faculties_content.full} into {target_restore_dir.full}"
                    )
                    backed_up_faculties_content.copy(
                        target_restore_dir.full, overwrite=False
                    )  # overwrite=False prefers destination (existing)
                else:
                    logger.warning(
                        f"Expected content for '{target_restore_dir.name}' not found in backup. Nothing to merge."
                    )
            else:  # Should be caught by Typer if strategy is an Enum
                logger.error(
                    f"Unrecognized restore strategy: {strategy}. Cannot restore backup."
                )
                return
            logger.info(
                f"Backup restored to {target_restore_dir.full} using strategy: {strategy.value}"
            )
        except Exception as e_restore:
            logger.error(f"Error during restore operation: {e_restore}")
            logger.debug(traceback.format_exc())
