import contextlib
import os
import pathlib
import shutil
import time
from datetime import datetime
import logging # Added

# Removed loguru and rich
# from loguru import logger
# from rich.console import Console

# # rich Console + overload the print function
# cons = Console(emoji=True, markup=True)
# print: callable = cons.print

logger = logging.getLogger(__name__) # Added

# Removed custom info, warn, cool functions
# def info(text: str) -> None:
#     logger.info(text)
#
#
# def warn(text: str) -> None:
#     logger.warning(text)
#
#
# def cool(text: str) -> None:
#     logger.success(text)


def determine_course_code(code: str, name: str) -> set[str]:
    """
    For a given course code and name (cols of a copyright item; canvas data), determine the correct osiris course code(s).
    Returns a set of string course codes; if no valid course code could be found it will be empty.
    Heuristic is as follows:

    STEP 1: attempt to parse canvas course code into osiris course code(s)
        - from column 'course_code', get the course code as a string
        - Should look like YYYY - XXXXXXXXXXX - 1A, where YYYY is the year, XXXXXXXXXXX is the course code, and 1A is the period.
        - split on '-', select the second part.
        - course code should be numeric and (probably?) 9 digits long.
        - period is (probably) one value from: JAAR, 1A, 1B, 2A, 2B, 3A, SEM1, SEM2, SEM3

        EXAMPLES:
            should result in extracted course code + period:
                2024-191158500-JAAR
                    --> Course code: 191158500, Period: JAAR
                    --> return {191158500}
                2024-201800005-1A
                    --> Course code: 201800005, Period: 1A
                    --> return {201800005}
                2024-202400157-1A
                    --> Course code: 202400157, Period: 1A
                    --> return {202400157}
                2024-201800236-SEM1
                    --> Course code: 201800236, Period: SEM1
                    --> return {201800236}

            should be processed further:
                2024-IDVWI-1A
                    --> Course code: IDVWI, Period: 1A
                    --> ERROR: not a valid course code
                    --> continue to step 2
                2024-ELECMSE-1B
                    --> Course code: ELECMSE, Period: 1B
                    --> ERROR: not a valid course code
                    --> continue to step 2

    If step 1 fails:
    STEP 2: attempt to parse canvas course name into osiris course code(s)
    in cases where the 'course code' is a string with only letters, it is likely this course has multiple course codes attached to it.
    in this case, the set of related course codes should be extracted from the 'course name' column.
        1. retrieve the string to parse from the 'course name' column.
        2. split the string on ';'. Split the second item of result on '(', select the first item of that result. This should give a string of course codes separated by commas.
        3. split on ',' and loop over results
        4. For each: if str with only digits and len >= 8: add to result set, set found to True.

        EXAMPLES:
            should result in extracted course codes:
                Circuit Analysis 1 and 2; 202001116,202200163 (2024-JAAR)
                    --> return {202001116, 202200163}
                Characterization of Nanostructures 2023; 193700010,201600043 (2024-1A)
                    --> return {193700010, 201600043}

            should not result in extracted course codes:
                Circuit Analysis 1 and 2; CA12,CA34 (2024-JAAR)
                    --> Course codes found: [CA12, CA34]
                    --> ERROR: invalid course codes
                    --> return empty list

    """

    def is_valid_course_code(check_code: Any) -> bool:
        """Checks if a given code is a valid Osiris course code (numeric, >= 8 digits)."""
        try:
            check_code_str = str(check_code).strip()

            return bool(check_code_str.isdigit() and len(check_code_str) >= 8)
        except Exception:
            return False

    tempresults: set[str] = set()
    first_try: str = ""
    second_try: str = ""

    try:
        first_try = code.split("-")[1].strip()
        tempresults.add(first_try) if is_valid_course_code(first_try) else None
        if (";" in name) and ("(" in name):
            second_try = name.split(";")[1].split("(")[0]
            [
                tempresults.add(c.strip())
                for c in second_try.split(",")
                if is_valid_course_code(c)
            ]

        if not tempresults:
            logger.warning(f"No valid course code found for {code} - {name}")
            logger.info( # Changed from utils.info to logger.info
                f"code extraction results: {first_try}, name extraction results: {second_try}"
            )
        return tempresults
    except Exception as e:
        logger.warning(f"Error in determine_course_code for input: code={code}, name={name}: {e}")
        return tempresults


# ----------------------------------------------------------------------------------------------------------------------
# Classes for handling files and directories.
# ----------------------------------------------------------------------------------------------------------------------


class Directory:
    """
    Simple class for directories + operations.

    Attributes:
        full (pathlib.Path): The full, absolute path to the directory.
        input_path_str (str): The string representation of the path provided during initialization.
        create_dir (bool): Whether the directory should be created if it doesn't exist upon initialization.
    """

    full: pathlib.Path
    input_path_str: str
    create_dir: bool

    def __init__(self, path: str | pathlib.Path, create_dir: bool = True) -> None:
        """
        Initializes a Directory object.

        Args:
            path (str | pathlib.Path): The path to the directory. Can be absolute or relative.
            create_dir (bool, optional): If True, creates the directory if it doesn't exist. Defaults to True.
        """
        self.input_path_str = str(path)
        resolved_path = pathlib.Path(path)
        self.create_dir = create_dir

        if resolved_path.is_absolute():
            self.full = resolved_path
        else:
            self.full = pathlib.Path.cwd() / resolved_path

        self._post_init() # Changed to internal call

    def _post_init(self) -> None:
        """
        Internal post-initialization hook.
        Checks if the path is a directory or creates it if create_dir is True.
        Raises NotADirectoryError if the path exists but is not a directory.
        """
        if not self.full.exists():
            if self.create_dir:
                self.create()
            else:
                logger.warning( # Changed from utils.warn to logger.warning
                    text=f"Directory {self.full} does not exist and create_dir is set to False. Call create() before any other commands!"
                )
        elif not self.full.is_dir():
            raise NotADirectoryError(f"Directory {self.full} is not a directory.")

    @property
    def files(self) -> list['File']:
        """Gets all files directly within this directory."""
        if not self.is_dir:
            return []
        return [
            File(path=item) for item in self.full.iterdir() if item.is_file()
        ]

    @property
    def files_r(self) -> list['File']:
        """Recursively gets all files within this directory and its subdirectories."""
        if not self.is_dir:
            return []
        return [
            File(path=item) for item in self.full.rglob("*") if item.is_file()
        ]

    @property
    def name(self) -> str:
        """The name of the directory (final part of the path)."""
        return self.full.name

    @property
    def created(self) -> datetime:
        """Timestamp of when the directory was created."""
        return datetime.fromtimestamp(self.full.stat().st_ctime) # Use fromtimestamp and st_ctime

    @property
    def dirs(self, r: bool = False) -> list['Directory']: # type: ignore # Pylance complains about @property with params
        """
        Returns a list of Directory objects within this directory.

        Args:
            r (bool, optional): If True, recursively finds all subdirectories. Defaults to False.

        Returns:
            list[Directory]: A list of Directory objects.
        """
        if not self.is_dir:
            return []
        if not r:
            return [
                Directory(path=item, create_dir=False)
                for item in self.full.iterdir()
                if item.is_dir()
            ]
        # r is True, get recursively
        return [
            Directory(path=item, create_dir=False)
            for item in self.full.rglob("*")
            if item.is_dir()
        ]

    def newest_file(self, file_type: list[str] | str | None = None) -> 'File | None':
        """
        Returns the newest file in the directory, optionally filtered by file type.

        Args:
            file_type (list[str] | str | None, optional): A file extension (e.g., ".txt")
                or list of extensions to filter by. Defaults to None (all files).

        Returns:
            File | None: The newest File object, or None if no matching files are found.
        """
        all_files: list[File] = self.files
        if file_type:
            extensions_to_check = [file_type] if isinstance(file_type, str) else file_type
            # Ensure extensions start with a dot for consistent comparison if needed, though Path.suffix includes it.
            extensions_to_check = [ext if ext.startswith('.') else f".{ext}" for ext in extensions_to_check]

            filtered_files = []
            for file_obj in all_files:
                # file.extension from File class already includes the dot
                if file_obj.extension in extensions_to_check:
                    # Original logic had `if "overview" not in file.name`. Keeping if still relevant.
                    if "overview" not in file_obj.name: # Assuming file.name is just the filename.ext
                        filtered_files.append(file_obj)
            all_files = filtered_files

        if not all_files:
            return None
        return max(all_files, key=lambda f: f.created)

    @property
    def newest_file_r(self) -> 'File | None': # Corrected return type from str
        """Recursively gets the newest file in the directory and its subdirectories."""
        all_files: list[File] = self.files_r
        if not all_files:
            return None
        return max(all_files, key=lambda f: f.created)

    @property
    def exists(self) -> bool:
        """Checks if the directory exists."""
        return self.full.exists()

    @property
    def is_dir(self) -> bool:
        """Checks if the path points to an existing directory."""
        return self.full.is_dir()

    def create(self) -> None:
        """Creates the directory, including any necessary parent directories."""
        with contextlib.suppress(FileExistsError): # Safely ignore if dir already exists
            self.full.mkdir(parents=True, exist_ok=True) # exist_ok=True is often more robust

    def delete(self) -> None:
        """Deletes the directory and all its contents recursively."""
        if self.exists and self.is_dir: # Ensure it exists and is a directory before deleting
            shutil.rmtree(self.full)
        elif self.exists and not self.is_dir:
            logger.warning(f"Path {self.full} is a file, not a directory. Cannot use rmtree.")
        else:
            logger.info(f"Directory {self.full} does not exist. Nothing to delete.")


    def copy(self, target: pathlib.Path | str, overwrite: bool = True) -> None:
        """
        Copies the directory to a new location.

        Args:
            target (pathlib.Path | str): The destination path.
            overwrite (bool, optional): If True, overwrites existing files at the destination.
                                       If False, only copies new files. Defaults to True.
        """
        def copy_only_new(src: str, dst: str, *, follow_symlinks: bool = True) -> str:
            if pathlib.Path(dst).exists():
                return dst # Skip if destination exists
            return shutil.copy2(src, dst, follow_symlinks=follow_symlinks)

        target_path = pathlib.Path(target) if isinstance(target, str) else target

        logger.info(f"Copying directory {self.full} to {target_path}")
        if not self.exists:
            logger.warning(f"Source directory {self.full} does not exist. Cannot copy.")
            return

        copy_func = shutil.copy2 if overwrite else copy_only_new
        shutil.copytree(
            src=self.full,
            dst=target_path,
            dirs_exist_ok=True, # Important for merging or overwriting parts of existing target
            copy_function=copy_func if not overwrite else None # copy_function is only used when dirs_exist_ok=True and we want custom file copy
        )


    def rename_latest_file(self, new_name: str) -> None:
        """Renames the newest file in this directory."""
        newest_file_obj = self.newest_file() # Renamed variable for clarity
        if newest_file_obj:
            newest_file_obj.rename(new_name)
        else:
            logger.warning(f"No files found in directory {self.full} to rename.")


    def __eq__(self, other: object) -> bool:
        """Checks if this Directory object is equal to another (based on full path)."""
        if isinstance(other, Directory):
            return self.full == other.full
        return False

    def __str__(self) -> str:
        """String representation of the Directory object (its full path)."""
        return str(object=self.full)

    def __repr__(self) -> str:
        return f"DirPath('{self.input_path_str}') -> {self.full}"


class File:
    """
    Simple class for files + operations.

    Attributes:
        _path (pathlib.Path): The internal Path object representing the file.
        _path_init_str (str): The initial string representation of the path.
        _name (str): The name of the file, including extension.
        _extension (str): The file extension.
        _dir (Directory): The Directory object representing the parent directory.
    """
    _path: pathlib.Path
    _path_init_str: str
    _name: str
    _extension: str
    _dir: Directory

    def __init__(self, path: str | pathlib.Path) -> None:
        """
        Initializes a File object.

        Args:
            path (str | pathlib.Path): The path to the file. Can be absolute or relative.
                                       Should include the filename and extension.

        Raises:
            TypeError: If the path is not a string or pathlib.Path.
        """
        self._path_init_str = str(path)

        if not isinstance(path, (str, pathlib.Path)):
            raise TypeError(f"Path must be a string or pathlib.Path, got {type(path)}")

        resolved_path = pathlib.Path(path)

        if not resolved_path.is_absolute():
            # If relative, resolve it based on CWD, then get parts
            # This ensures _dir is always an absolute path Directory object
            resolved_path = pathlib.Path.cwd() / resolved_path

        self._path = resolved_path
        self._name = resolved_path.name
        self._extension = resolved_path.suffix
        # Parent directory should also be resolved to an absolute path
        self._dir = Directory(path=resolved_path.parent, create_dir=False) # create_dir=False as parent should exist or it's an issue with path


    @property
    def exists(self) -> bool:
        """Checks if the file exists."""
        return self._path.exists() and self._path.is_file() # Ensure it's a file too

    @property
    def is_file(self) -> bool:
        """Checks if the path points to an existing file."""
        return self._path.is_file()

    @property
    def path(self) -> pathlib.Path:
        """The full pathlib.Path object for the file."""
        return self._path

    @property
    def name(self) -> str:
        """The name of the file, including extension."""
        return self._name

    @property
    def extension(self) -> str:
        """The file extension (e.g., '.txt')."""
        return self._extension

    @property
    def dir(self) -> Directory:
        """The parent Directory object."""
        return self._dir

    @property
    def created(self) -> datetime:
        """Timestamp of when the file was created (st_birthtime or st_ctime)."""
        # st_birthtime is not available on all systems (e.g. some Linux)
        # st_ctime is the last metadata change time on Unix, or creation time on Windows.
        try:
            return datetime.fromtimestamp(self._path.stat().st_birthtime)
        except AttributeError: # pragma: no cover
            return datetime.fromtimestamp(self._path.stat().st_ctime)


    @property
    def modified(self) -> datetime:
        """Timestamp of when the file was last modified."""
        return datetime.fromtimestamp(self._path.stat().st_mtime)

    @property
    def size(self) -> int:
        """Size of the file in bytes."""
        return self._path.stat().st_size

    def copy(self, new_path: str | pathlib.Path) -> 'File':
        """
        Copies the file to a new location.

        Args:
            new_path (str | pathlib.Path): The destination path for the copy.

        Returns:
            File: A new File object representing the copied file.
        """
        target_path = pathlib.Path(new_path) if isinstance(new_path, str) else new_path
        if not self.exists:
            logger.warning(f"Source file {self._path} does not exist. Cannot copy.")
            # Depending on desired strictness, could raise FileNotFoundError
            return File(path=target_path) # Return a File object for the target path anyway

        shutil.copy(src=self._path, dst=target_path)
        return File(path=target_path)

    def move(self, new_path: str | pathlib.Path) -> 'File':
        """
        Moves the file to a new location.
        If the target file already exists, it appends a timestamp to the new filename to avoid overwrite by default.

        Args:
            new_path (str | pathlib.Path): The destination path. This can be a directory (file will be moved into it with original name)
                                         or a full file path (file will be moved/renamed to this path).

        Returns:
            File: A new File object representing the moved file at its new location.
        """
        target_path = pathlib.Path(new_path) if isinstance(new_path, str) else new_path

        if not self.exists:
            logger.warning(f"Source file {self._path} does not exist. Cannot move.")
            return File(path=target_path)


        # If target_path is an existing directory, move the file into it with its original name
        if target_path.is_dir():
            final_target_path = target_path / self.name
        else: # Assume target_path is a full file path (or a non-existent path where the file should be placed)
            final_target_path = target_path
            # Ensure parent directory of the target file path exists
            final_target_path.parent.mkdir(parents=True, exist_ok=True)


        if final_target_path.exists():
            timestamp_suffix = f"_{int(time.time())}"
            final_name_versioned = f"{final_target_path.stem}{timestamp_suffix}{final_target_path.suffix}"
            final_target_path = final_target_path.with_name(final_name_versioned)
            logger.info(f"Target {new_path} exists or would overwrite. Moving to versioned path: {final_target_path}")


        shutil.move(src=str(self._path), dst=str(final_target_path.resolve()))
        return File(path=final_target_path)

    def rename(self, new_name: str) -> 'File':
        """
        Renames the file within its current directory.

        Args:
            new_name (str): The new file name (including extension).

        Returns:
            File: The same File object, now representing the renamed file.

        Raises:
            FileNotFoundError: If the original file does not exist.
        """
        if not self.exists:
            raise FileNotFoundError(f"File {self._path} does not exist, cannot rename.")

        new_path = self._dir.full / new_name
        self._path = self._path.rename(new_path)
        self._name = new_name
        self._extension = new_path.suffix # Update extension if new_name changed it
        return self

    def delete(self) -> None:
        """Deletes the file. Does not raise an error if the file is already missing."""
        self._path.unlink(missing_ok=True)

    def __eq__(self, other: object) -> bool:
        """Checks if this File object is equal to another (based on resolved absolute path)."""
        if isinstance(other, File):
            return self._path.resolve() == other._path.resolve()
        if isinstance(other, str): # Allow comparison with string path
            return self._path.resolve() == pathlib.Path(other).resolve()
        return False

    def __str__(self) -> str:
        """String representation of the File object (its full path)."""
        return str(self._path)

    def __repr__(self) -> str:
        """Detailed string representation of the File object."""
        return str(object=self._path.absolute())
