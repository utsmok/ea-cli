import pathlib
import os
import shutil
from datetime import datetime
import time
from rich.console import Console
from loguru import logger

# rich Console + overload the print function
cons = Console(emoji=True, markup=True)
print: callable = cons.print

def info(text: str) -> None:
    logger.info(text)

def warn(text: str) -> None:
    logger.warning(text)

def cool(text: str) -> None:
    logger.success(text)



def determine_course_code(code: str, name: str) -> set[str|None]:
    """
    For a given course code and name (cols of a copyright item; canvas data), determine the correct osiris course code(s).
    Returns a set of course codes; if no valid course code could be found it will be empty.
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
    def is_valid_course_code(check_code) -> bool:
        try:
            check_code = str(check_code).strip()

            if check_code.isdigit() and len(check_code) >= 8:
                return True
            else:
                return False
        except Exception as e:
            return False

    tempresults = set()
    first_try = ""
    second_try = ""

    try:
        first_try = code.split("-")[1].strip()
        tempresults.add(first_try) if is_valid_course_code(first_try) else None
        if (';' in name) and ('(' in name):
            second_try = name.split(";")[1].split("(")[0]
            [tempresults.add(c.strip()) for c in second_try.split(",") if is_valid_course_code(c)]

        if not tempresults:
            warn(f"No valid course code found for {code} - {name}")
            info(
                f"code extraction results: {first_try}, name extraction results: {second_try}"
            )
        return tempresults
    except Exception as e:
        warn(f"Error in determine_course_code for input: code={code}, name={name}: {e}")
        return tempresults

# ----------------------------------------------------------------------------------------------------------------------
# Classes for handling files and directories.
# ----------------------------------------------------------------------------------------------------------------------

class Directory:
    """
    Simple class for directories + operations.
    Init with an absolute path, or a path relative to the current working directory.
    If the dir does not yet exist, it will be created. Disable this by setting the 'create_dir' parameter to False.
    """
    full: pathlib.Path
    def __init__(self, path: str | pathlib.Path, create_dir: bool = True) -> None:
        if isinstance(path, pathlib.Path):
            self.input_path_str = str(object=path)
        else:
            self.input_path_str = path
            path = pathlib.Path(path)
        self.create_dir: bool = create_dir

        # check if the path is absolute
        if path.is_absolute():
            self.full = path
        else:
            self.full = pathlib.Path.cwd() / path

        self.post_init()

    def post_init(self) -> None:
        """
        Checks to see if this is actually a dir,
        or create it if create_dir is set to True.
        """
        if not self.full.exists():
            if self.create_dir:
                self.create()
            else:
                warn(
                    text=f"Directory {self.full} does not exist and create_dir is set to False. Call create() before any other commands!"
                )
        elif not self.full.is_dir():
            raise NotADirectoryError(f"Directory {self.full} is not a directory.")

    @property
    def files(self) -> list["File"]:
        """
        Gets all files in the dir as a list of File objects.
        """
        return [
            File(path=self.full / file) for file in self.full.iterdir() if file.is_file()
        ]

    @property
    def files_r(self) -> list["File"]:
        """
        Recursively gets all files in the dir, so including files in subdirs, as a list of File objects.
        """
        return [
            File(path=self.full / file) for file in self.full.rglob(pattern="*") if file.is_file()
        ]
    @property
    def name(self) -> str:
        return self.full.name
    @property
    def created(self) -> datetime:
        return datetime.strptime(time.ctime(self.full.stat().st_birthtime), "%c")
    @property
    def dirs(self, r: bool = False) -> list["Directory"]:
        """
        Returns a list of all dirs in this Directory as a list of Directory objects.
        If r is set to True, it will return all children dirs recursively.
        """
        if not r:
            return [Directory(path=str(d), create_dir=False) for d in self.full.iterdir() if d.is_dir()]
        if r:
            return [Directory(path=str(d), create_dir=False) for d in self.full.rglob(pattern="*") if d.is_dir()]

    def newest_file(self, file_type:list[str]|str|None = None) -> "File":
        """
        Returns the newest file in the dir as a File object.
        Parameters:
            file_type (str): If set, only files with this extension will be returned.
            input the extension incl dot; or a list of them.
        """
        all_files: list[File] = self.files
        if file_type:
            if isinstance(file_type, str):
                file_type = [file_type]
            all_files = [file for file in all_files if file.extension in file_type if 'overview' not in file.name]
        if not all_files:
            return None
        return max(all_files, key=lambda x: x.created)


    @property
    def newest_file_r(self) -> str:
        """
        Recursively gets the newest file in the dir, so including files in subdirs, as a File object.
        """
        all_files: list[File] = self.files_r
        return max(all_files, key=lambda x: x.created)

    @property
    def exists(self) -> bool:
        return self.full.exists()

    @property
    def is_dir(self) -> bool:
        return self.full.is_dir()

    def create(self) -> None:
        try:
            self.full.mkdir(parents=True, exist_ok=False)
        except FileExistsError:
            pass

    def delete(self) -> None:
        shutil.rmtree(self.full)

    def copy(self, target: pathlib.Path | str, overwrite:bool = True) -> None:
        def copy_only_new(src, dst, *, follow_symlinks=True):
            if dst.exists():
                return dst
            else:
                return shutil.copy2(src=src, dst=dst, follow_symlinks=follow_symlinks)

        if isinstance(target, Directory):
            target = target.full

        info(text=f"Copying {self.full} to {target}")
        if not overwrite:
            info(text='Overwrite set to False, copying only new files.')
            shutil.copytree(src=self.full, dst=target, dirs_exist_ok=True, copy_function=copy_only_new)
        else:
            info(text='Overwrite set to True, copying all files.')
            shutil.copytree(src=self.full, dst=target, dirs_exist_ok=True)

    def rename_latest_file(self, new_name: str) -> None:
        newest_file = self.newest_file()
        if newest_file:
            newest_file.rename(new_name)

    def __eq__(self, other) -> bool:
        return self.full == other.full

    def __str__(self) -> str:
        return str(object=self.full)

    def __repr__(self) -> str:
        return f"DirPath('{self.input_path_str}') -> {self.full}"

class File:
    """
    Simple class for files + operations
    Parameters:
        path: str or Path
            relative from the current working directory.
            OR
            absolute path to the file.
            Should always end with the filename including extension.
    """

    def __init__(self, path: str | pathlib.Path) -> None:
        self._path_init_str = str(object=path)

        assert isinstance(path, str) or isinstance(path, pathlib.Path)

        if isinstance(path, pathlib.Path):
            self._path: pathlib.Path = path
            self._name: str = path.name
            self._extension: str = path.suffix
            self._dir = Directory(path=str(object=self._path.absolute().parent))
        elif isinstance(path, str):
            if "/" in path:
                self._name = path.rsplit(sep="/", maxsplit=1)[-1]
                self._dir = Directory(path=path.rsplit(sep="/", maxsplit=1)[0], create_dir=True)
            else:
                self._name = path
                self._dir = Directory(path=os.getcwd())

            self._extension = self._name.split(sep=".")[-1]
            self._path = self._dir.full / self._name

    @property
    def exists(self) -> bool:
        return self._path.exists()

    @property
    def is_file(self) -> bool:
        return self._path.is_file()

    @property
    def path(self) -> pathlib.Path:
        return self._path

    @property
    def name(self) -> str:
        return self._name

    @property
    def extension(self) -> str:
        return self._extension

    @property
    def dir(self) -> Directory:
        return self._dir

    @property
    def created(self) -> datetime:
        return datetime.fromtimestamp(timestamp=self._path.stat().st_birthtime)

    @property
    def modified(self) -> datetime:
        return datetime.fromtimestamp(timestamp=self._path.stat().st_mtime)

    @property
    def size(self) -> int:
        return self._path.stat().st_size
    def copy(self, new_path: str) -> "File":
        shutil.copy(src=self._path, dst=new_path)
        return File(path=new_path)

    def move(self, new_path: str | pathlib.Path) -> "File":
        '''
        Moves the file to the indicated new path.
        If the
        '''
        if isinstance(new_path, str):
            new_path = pathlib.Path(new_path)
        if new_path.exists():
            if '.' in str(object=new_path):
                new_path = pathlib.Path(str(object=new_path).split(sep=".")[0]+"_" +str(object=int(time.time()))+'.'+str(object=new_path).split(sep=".")[1])
            else:
                new_path = pathlib.Path(str(object=new_path)+"_" +str(object=int(x=time.time())))
        shutil.move(src=self._path, dst=str(object=new_path.absolute()))
        return File(path=new_path)

    def rename(self, new_name: str) -> "File":
        os.rename(src=self._path, dst=self._dir.full / new_name)
        self._path = self._dir.full / new_name

    def delete(self) -> None:
        os.remove(path=self._path)

    def __eq__(self, other: "File") -> bool:
        if isinstance(other, str):
            return any([self._path == pathlib.Path(other), self._name == other])
        return self._path == other.path

    def __str__(self) -> str:
        return str(object=self._path)

    def __repr__(self) -> str:
        return str(object=self._path.absolute())
