from easy_access.utils import Directory, cool, warn, info, print
from easy_access.settings import SETTINGS, DirSetting
from datetime import datetime
from enum import Enum

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
        """
        This class contains the functions for handling and restoring backups.
        Restoring is handled with cli flags, see run.py in the root dir (use --help for more details in your cli).
        Backup settings are in the 'backup' and 'directories' sections of the settings.yaml file:
        <settings.yaml>
            backup:
                backup_all: true  #backup all sheets in selected directories before starting? (bool)
                backup_dirs: #directories to backup, use keys from the 'directories' section in settings.yaml
                    - faculties_dir
                max_backups: 3                 # maximum number of backups to keep (int)
            directories:
                full_backups: full_backups    # directory to store the backups (str - relative path)
        """
        ...

    def backup_files(self) -> None:
        dirs_to_backup = SETTINGS.backup_settings.backup_dirs
        backup_location = SETTINGS.backup_settings.backup_location
        max_backups = SETTINGS.backup_settings.max_backups

        if not dirs_to_backup:
            warn('backup_all is set to true in settings.yaml, but no dirs to backup were specified. Skipping.')
            return
        if not backup_location:
            warn('backup_all is set to true in settings.yaml, but no backup location was specified. Skipping.')
            return
        if not max_backups:
            warn('backup_all is set to true in settings.yaml, but no max amount of backups was specified. Skipping.')
            return

        backup_subdirs = backup_location.dirs
        if len(backup_subdirs) > max_backups:
            while len(backup_subdirs) > max_backups:
                min(backup_subdirs, key=lambda x: x.created).delete()
        info(f"Creating backup of all data in dirs: {[d.full.name for d in dirs_to_backup]}")
        for i in range(0, max_backups + 4):
            new_backup_dir = Directory(backup_location.full / f"backup{i if i > 0 else ""}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}", create_dir=False)
            if not new_backup_dir.exists:
                new_backup_dir.create()
                break

        for d in dirs_to_backup:
            d.copy(new_backup_dir)

        cool(f'Backups done, stored in {new_backup_dir.full}')

    def restore_backup(self, backup_dir: Directory = None, strategy: RestoreStrategy = RestoreStrategy.REPLACE, select: RestoreOptions = RestoreOptions.LATEST) -> None:
        """
        Restore a backup from a directory. Defaults to restoring the latest backup.

        Parameters:
            backup_dir: Directory: the directory to restore from. If not specified, the dir in backup_settings will be used.
            strategy: RestoreStrategy: the strategy to use for restoring the backup. options:
                                - "replace" (default): remove 'faculties' dir and replace with backup dir
                                - "merge_prefer_backup": merge the backup dir with the 'faculties' dir: overwrite files with the same name, add new files, keep old files
                                - "merge_prefer_existing": merge the backup dir with the 'faculties' dir: DO NOT overwrite files that already exist. Add new files, keep existing files
            select: RestoreOptions: Which backup to restore. Valid options:
                                - "latest" (default): restore the latest backup
                                - "oldest": restore the oldest backup
                                - "manual": let the user select which backup to restore
        """

        target_dir = SETTINGS.dirs[DirSetting.FACULTIES_DIR]

        if not backup_dir:
            backup_dir = SETTINGS.backup_settings.backup_location
        if not isinstance(backup_dir, Directory):
            try:
                backup_dir = Directory(backup_dir)
            except Exception:
                warn(f"Could not convert {backup_dir} to a Directory object. Cannot restore backup.")
                return
        if not backup_dir.exists:
            warn(f"Backup dir {backup_dir.full} does not exist -- cannot restore backup.")
            return

        selected_backup_dir = None
        if select == RestoreOptions.LATEST:
            selected_backup_dir = max(backup_dir.dirs, key=lambda x: x.created)
        elif select == RestoreOptions.OLDEST:
            selected_backup_dir = min(backup_dir.dirs, key=lambda x: x.created)
        elif select == RestoreOptions.MANUAL:
            print(f"  Select backup to restore from {backup_dir.full}:")
            print("------------------------------------------------------\n")
            for num, dir in enumerate(backup_dir.dirs):
                print(f"    {num}: {dir.full}")
            print("\n")
            while not selected_backup_dir:
                try:
                    selected_backup_dir = backup_dir.dirs[int(input(f"  Select backup to restore (0-{len(backup_dir.dirs) - 1}): "))]
                except Exception as e:
                    warn(f"Invalid input. Please select a valid backup. ({e})")

        if not selected_backup_dir:
            warn("No valid backup selected, cannot restore.")
            return

        if strategy == RestoreStrategy.REPLACE:
            target_dir.delete()
            selected_backup_dir.copy(target_dir)
        elif strategy == RestoreStrategy.MERGE_PREFER_BACKUP:
            selected_backup_dir.copy(target_dir)
        elif strategy == RestoreStrategy.MERGE_PREFER_EXISTING:
            target_dir.copy(selected_backup_dir, overwrite=False)
        else:
            warn(f"Unrecognized strategy: {strategy}. Cannot restore backup.")
            return
        cool(f"Backup restored to {target_dir.full} using strategy: {strategy}")
