from __future__ import annotations

import contextlib
import os
import pathlib
import shutil
import time
from collections.abc import Callable
from datetime import datetime
from typing import Any

from loguru import logger

def print(text:str):
    print(text)
    logger.warning('a print function was called...')

def determine_course_code(code: str, name: str) -> set[str]:
    """Determines Osiris course code(s) from Canvas course data.

    This function attempts to parse the correct Osiris course code(s) from
    the provided Canvas course code and name.

    The heuristic is as follows:

    STEP 1: Attempt to parse Canvas course code into Osiris course code(s).
        - From the 'course_code' string (e.g., "YYYY - XXXXXXXXXXX - 1A").
        - Extract the middle part as the potential course code.
        - A valid Osiris course code is typically numeric and around 9 digits long.

        Examples:
            - "2024-191158500-JAAR" -> "191158500"
            - "2024-201800005-1A" -> "201800005"

    STEP 2: If Step 1 fails, attempt to parse Canvas course name.
        - This is used when the 'course_code' field contains non-numeric values
          (e.g., "2024-IDVWI-1A"), suggesting multiple codes might be in the name.
        - The 'course_name' might look like:
          "Circuit Analysis 1 and 2; 202001116,202200163 (2024-JAAR)"
        - Extract codes from the part after ';' and before '('.
        - Split by ',' and validate each part.

        Examples:
            - "Circuit Analysis 1 and 2; 202001116,202200163 (2024-JAAR)"
              -> {"202001116", "202200163"}
            - "Characterization of Nanostructures 2023; 193700010,201600043 (2024-1A)"
              -> {"193700010", "201600043"}

    Args:
        code: The course code string from Canvas data.
        name: The course name string from Canvas data.

    Returns:
        A set of valid Osiris course codes found. Returns an empty set if
        no valid codes could be determined.
    """

    def is_valid_course_code(check_code: Any) -> bool:
        """Checks if a given string is a plausible Osiris course code."""
        try:
            check_code_str = str(check_code).strip()
            return bool(check_code_str.isdigit() and len(check_code_str) >= 8)
        except Exception:
            return False

    temp_results: set[str] = set()
    first_try_code: str = ""
    second_try_codes_str: str = ""

    try:
        # Step 1: Attempt to parse from 'code'
        parts = code.split("-")
        if len(parts) > 1:
            first_try_code = parts[1].strip()
            if is_valid_course_code(first_try_code):
                temp_results.add(first_try_code)

        # Step 2: Attempt to parse from 'name' if necessary or if name contains codes
        if ";" in name and "(" in name:
            try:
                name_parts = name.split(";", 1)
                if len(name_parts) > 1:
                    codes_section = name_parts[1].split("(", 1)[0]
                    second_try_codes_str = codes_section.strip()
                    for c in second_try_codes_str.split(","):
                        c_stripped = c.strip()
                        if is_valid_course_code(c_stripped):
                            temp_results.add(c_stripped)
            except IndexError:  # Handle cases where splitting might fail
                pass

        if not temp_results:
            logger.warning(f"No valid course code found for {code} - {name}")
            logger.info(
                f"Code extraction attempt: '{first_try_code}', Name extraction attempt: '{second_try_codes_str}'"
            )
        return temp_results
    except Exception as e:
        logger.warning(
            f"Error in determine_course_code for input: code='{code}', name='{name}': {e}"
        )
        return temp_results


class Directory:
    """Represents a directory and provides operations on it.

    Attributes:
        full (pathlib.Path): The full, absolute path to the directory.
        input_path_str (str): The original path string provided during initialization.
        create_dir (bool): Whether the directory should be created if it doesn't exist.
    """

    full: pathlib.Path
    input_path_str: str
    create_dir: bool

    def __init__(self, path: str | pathlib.Path, create_dir: bool = True) -> None:
        """Initializes a Directory object.

        Args:
            path (str | pathlib.Path): The path to the directory (absolute or relative to CWD).
            create_dir (bool) (default=True): If True, creates the directory if it doesn't exist.
        """
        if isinstance(path, Directory):
            # If a Directory object is passed (which shouldn't happen), use its full path and move on
            path = path.full

        self.input_path_str = str(path)
        self.create_dir = create_dir

        if isinstance(path, str):
            path = pathlib.Path(path)

        if not path.is_absolute():
            self.full = (pathlib.Path.cwd() / path).resolve()
        else:
            self.full = path.resolve()

        self._post_init()

    def _post_init(self) -> None:
        """Performs post-initialization checks and directory creation."""
        if not self.full.exists():
            if self.create_dir:
                self.create()
            else:
                logger.warning(
                    f"Directory {self.full} does not exist and create_dir is False. "
                    "Call create() before other commands."
                )
        elif not self.full.is_dir():
            raise NotADirectoryError(f"Path {self.full} exists but is not a directory.")

    @property
    def files(self) -> list[File]:
        """list[File]: All files directly within this directory."""
        if not self.is_dir:
            return []
        return [
            File(file_path) for file_path in self.full.iterdir() if file_path.is_file()
        ]

    @property
    def files_r(self) -> list[File]:
        """list[File]: All files in this directory and its subdirectories (recursively)."""
        if not self.is_dir:
            return []
        return [
            File(file_path) for file_path in self.full.rglob("*") if file_path.is_file()
        ]

    @property
    def name(self) -> str:
        """str: The name of the directory."""
        return self.full.name

    @property
    def created(self) -> datetime:
        """datetime | None: The creation timestamp of the directory.

        Note:
            Uses `st_birthtime`, which may not be available on all platforms.
            Returns None if the timestamp cannot be retrieved or directory doesn't exist.
        """
        if not self.exists:
            raise FileNotFoundError(
                f"Directory {self.full} does not exist. Cannot retrieve creation time."
            )
        try:
            return datetime.fromtimestamp(self.full.stat().st_birthtime)
        except AttributeError:
            # st_birthtime might not be available, try st_ctime as a fallback
            try:
                return datetime.fromtimestamp(self.full.stat().st_ctime)
            except Exception as e:
                raise Exception(
                    f"Failed to retrieve creation time for {self.full}. Error: {e}"
                )
        except FileNotFoundError:
            raise FileNotFoundError(
                f"Directory {self.full} does not exist. Cannot retrieve creation time."
            )

    def dirs(self, r: bool = False) -> list[Directory]:
        """Gets subdirectories within this directory.

        Args:
            r: If True, recursively gets all subdirectories.

        Returns:
            A list of Directory objects.
        """
        if not self.is_dir:
            return []
        if r:
            return [
                Directory(d, create_dir=False)
                for d in self.full.rglob("*")
                if d.is_dir()
            ]
        else:
            return [
                Directory(d, create_dir=False)
                for d in self.full.iterdir()
                if d.is_dir()
            ]

    def newest_file(self, file_type: list[str] | str | None = None) -> File | None:
        """Gets the newest file in the directory, optionally filtered by type.

        Args:
            file_type: A file extension (e.g., ".txt") or a list of extensions
                       to filter by. Includes the dot.

        Returns:
            The newest File object, or None if no matching files are found.
        """
        if not self.is_dir:
            return None

        candidate_files: list[File] = self.files
        if file_type:
            if isinstance(file_type, str):
                extensions_to_check = [file_type.lower()]
            else:
                extensions_to_check = [ft.lower() for ft in file_type]

            candidate_files = [
                f
                for f in candidate_files
                if f.extension.lower() in extensions_to_check
                and "overview" not in f.name.lower()
            ]

        if not candidate_files:
            return None

        # Filter out files for which 'created' is None
        valid_files = [f for f in candidate_files if f.created is not None]
        if not valid_files:
            return None

        # Ensure the key for max always returns a comparable datetime object.
        # The 'else datetime.min' should not be hit due to prior filtering.
        return max(
            valid_files,
            key=lambda f: f.created if f.created is not None else datetime.min,
        )

    @property
    def newest_file_r(self) -> File | None:
        """File | None: The newest file in this directory or subdirectories (recursively).

        Returns None if no files are found or creation times are unavailable.
        """
        if not self.is_dir:
            return None
        all_files_recursive: list[File] = self.files_r
        if not all_files_recursive:
            return None

        valid_files = [f for f in all_files_recursive if f.created is not None]
        if not valid_files:
            return None

        # Ensure the key for max always returns a comparable datetime object.
        # The 'else datetime.min' should not be hit due to prior filtering.
        return max(
            valid_files,
            key=lambda f: f.created if f.created is not None else datetime.min,
        )

    @property
    def exists(self) -> bool:
        """bool: True if the directory exists, False otherwise."""
        return self.full.exists()

    @property
    def is_dir(self) -> bool:
        """bool: True if the path points to an existing directory."""
        return self.full.is_dir()

    def create(self) -> None:
        """Creates the directory, including any necessary parent directories."""
        with contextlib.suppress(FileExistsError):
            self.full.mkdir(parents=True, exist_ok=True)  # exist_ok=True is safer

    def delete(self) -> None:
        """Deletes the directory and all its contents recursively."""
        if self.exists:
            shutil.rmtree(self.full)

    def copy(
        self, target: str | pathlib.Path | Directory, overwrite: bool = True
    ) -> None:
        """Copies the directory to a target location.

        Args:
            target: The destination path or Directory object.
            overwrite: If True, overwrites existing files. If False, only new files are copied.
        """
        if not self.is_dir:
            logger.warning(
                f"Source directory {self.full} does not exist or is not a directory. Cannot copy."
            )
            return

        if isinstance(target, Directory):
            target_path = target.full
        elif isinstance(target, str):
            target_path = pathlib.Path(target)
        else:  # pathlib.Path
            target_path = target

        target_path = target_path.resolve()  # Ensure target is absolute

        def copy_only_new(src: str, dst: str, *, follow_symlinks: bool = True) -> str:
            """Copies only if destination does not exist."""
            if pathlib.Path(dst).exists():
                return dst
            return shutil.copy2(src, dst, follow_symlinks=follow_symlinks)

        logger.info(f"Copying {self.full} to {target_path}")
        if not overwrite:
            logger.info("Overwrite is False, copying only new files.")
            shutil.copytree(
                src=self.full,
                dst=target_path,
                dirs_exist_ok=True,
                copy_function=copy_only_new,
            )
        else:
            logger.info("Overwrite is True, copying all files.")
            shutil.copytree(src=self.full, dst=target_path, dirs_exist_ok=True)

    def rename_latest_file(self, new_name: str) -> None:
        """Renames the newest file in this directory.

        Args:
            new_name: The new name for the file (including extension).
        """
        latest_file = self.newest_file()
        if latest_file:
            latest_file.rename(new_name)
        else:
            logger.warning(f"No file found in {self.full} to rename.")

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Directory):
            return NotImplemented
        return self.full == other.full

    def __str__(self) -> str:
        return str(self.full)

    def __repr__(self) -> str:
        return f"Directory(path='{self.input_path_str}') -> {self.full}"


class File:
    """Represents a file and provides operations on it.

    Attributes:
        _path_init_str (str): The original path string provided during initialization.
        _path (pathlib.Path): The full, absolute path to the file.
        _name (str): The name of the file, including extension.
        _extension (str): The file extension (e.g., ".txt").
        _dir (Directory): The Directory object representing the file's parent directory.
    """

    _path_init_str: str
    _path: pathlib.Path
    _name: str
    _extension: str
    _dir: Directory

    def __init__(
        self, path: str | pathlib.Path, create_parent_dir: bool = False
    ) -> None:
        """Initializes a File object.

        Args:
            path: The path to the file (absolute or relative to CWD).
            create_parent_dir: If True, the parent directory will be created
                               if it doesn't exist. Defaults to False.
        """
        self._path_init_str = str(path)

        if isinstance(path, str):
            path = pathlib.Path(path)

        if not path.is_absolute():
            self._path = (pathlib.Path.cwd() / path).resolve()
        else:
            self._path = path.resolve()

        self._name = self._path.name
        self._extension = self._path.suffix
        self._dir = Directory(self._path.parent, create_dir=create_parent_dir)

    @property
    def exists(self) -> bool:
        """bool: True if the file exists, False otherwise."""
        return self._path.exists() and self._path.is_file()

    @property
    def is_file(self) -> bool:
        """bool: True if the path points to an existing file."""
        return self._path.is_file()

    @property
    def path(self) -> pathlib.Path:
        """pathlib.Path: The full, absolute path to the file."""
        return self._path

    @property
    def name(self) -> str:
        """str: The name of the file, including extension."""
        return self._name

    @property
    def extension(self) -> str:
        """str: The file extension (e.g., ".txt")."""
        return self._extension

    @property
    def dir(self) -> Directory:
        """Directory: The Directory object for the file's parent directory."""
        return self._dir

    @property
    def created(self) -> datetime:
        """datetime | None: The creation timestamp of the file.

        Note:
            Uses `st_birthtime`, which may not be available on all platforms.
            Returns None if timestamp cannot be retrieved or file doesn't exist.
        """
        if not self.exists:
            raise FileNotFoundError(
                f"File {self._path} does not exist. Cannot retrieve creation time."
            )
        try:
            return datetime.fromtimestamp(self._path.stat().st_birthtime)
        except AttributeError:
            # st_birthtime might not be available, try st_ctime as a fallback
            try:
                return datetime.fromtimestamp(self._path.stat().st_ctime)
            except Exception as e:
                raise Exception(
                    f"Failed to retrieve creation time for {self._path}: {e}"
                )

        except FileNotFoundError:
            raise FileNotFoundError(
                f"File {self._path} does not exist. Cannot retrieve creation time."
            )


    @property
    def modified(self) -> datetime | None:
        """datetime | None: The last modification timestamp of the file.

        Returns None if file doesn't exist.
        """
        if not self.exists:
            return None
        try:
            return datetime.fromtimestamp(self._path.stat().st_mtime)
        except FileNotFoundError:  # pragma: no cover
            return None

    @property
    def size(self) -> int | None:
        """int | None: The size of the file in bytes.

        Returns None if file doesn't exist.
        """
        if not self.exists:
            return None
        try:
            return self._path.stat().st_size
        except FileNotFoundError:  # pragma: no cover
            return None

    def copy(self, new_path: str | pathlib.Path) -> File | None:
        """Copies the file to a new path.

        Args:
            new_path: The destination path for the copy.

        Returns:
            A new File object for the copied file, or None if copy fails.
        """
        if not self.exists:
            logger.warning(f"File {self._path} does not exist. Cannot copy.")
            return None
        try:
            shutil.copy2(self._path, new_path)  # copy2 preserves more metadata
            return File(new_path)
        except Exception as e:
            logger.warning(f"Failed to copy {self._path} to {new_path}: {e}")
            return None

    def move(self, new_path: str | pathlib.Path) -> File | None:
        """Moves the file to a new path.

        If the target path exists, a timestamp is appended to the new filename
        to avoid overwriting.

        Args:
            new_path: The destination path.

        Returns:
            A File object representing the moved file at its new location,
            or None if move fails.
        """
        if not self.exists:
            logger.warning(f"File {self._path} does not exist. Cannot move.")
            return None

        target_path = pathlib.Path(new_path) if isinstance(new_path, str) else new_path
        target_path = target_path.resolve()  # Ensure target is absolute

        if target_path.exists():
            timestamp = str(int(time.time()))
            if target_path.is_dir():  # Moving into a directory
                target_path = (
                    target_path / f"{self._path.stem}_{timestamp}{self._path.suffix}"
                )
            else:  # Target is a file, append timestamp before suffix
                target_path = target_path.with_name(
                    f"{target_path.stem}_{timestamp}{target_path.suffix}"
                )

        try:
            moved_path_str = shutil.move(str(self._path), str(target_path))
            # Update self to reflect the move, or create a new object?
            # The original returned a new File object. Let's stick to that.
            return File(moved_path_str)
        except Exception as e:
            logger.warning(f"Failed to move {self._path} to {target_path}: {e}")
            return None

    def rename(self, new_name: str) -> File | None:
        """Renames the file within its current directory.

        Args:
            new_name: The new name for the file (e.g., "new_name.txt").

        Returns:
            The File object itself (now representing the renamed file),
            or None if rename fails.
        """
        if not self.exists:
            logger.warning(f"File {self._path} does not exist. Cannot rename.")
            return None

        new_path_target = self._dir.full / new_name
        try:
            # os.rename can be problematic across filesystems, pathlib.Path.rename is better
            renamed_path = self._path.rename(new_path_target)
            self._path = renamed_path  # Update internal path
            self._name = renamed_path.name
            self._extension = renamed_path.suffix
            # _path_init_str might be stale, but it's for initial reference
            return self
        except Exception as e:
            logger.warning(f"Failed to rename {self._path} to {new_path_target}: {e}")
            return None

    def delete(self) -> None:
        """Deletes the file."""
        if self.exists:
            try:
                os.remove(self._path)
            except Exception as e:  # pragma: no cover
                logger.warning(f"Failed to delete file {self._path}: {e}")
        else:
            logger.warning(f"File {self._path} does not exist. Cannot delete.")

    def __eq__(self, other: object) -> bool:
        if isinstance(other, File):
            return self._path == other._path
        if isinstance(other, pathlib.Path):
            return self._path == other.resolve()
        if isinstance(other, str):
            try:
                return self._path == pathlib.Path(other).resolve()
            except Exception:  # Handle invalid path strings
                return False
        return NotImplemented

    def __str__(self) -> str:
        return str(self.path)

    def __repr__(self) -> str:
        # Use resolve() for a canonical representation if path might be relative initially
        return f"File(path='{self._path_init_str}') -> {self._path.resolve()}"
