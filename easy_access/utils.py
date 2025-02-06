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

def info(text: str):
    logger.info(text)

def warn(text: str):
    logger.warning(text)

def cool(text: str):
    logger.success(text)


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
    def __init__(self, path: str | pathlib.Path, create_dir: bool = True):
        if isinstance(path, pathlib.Path):
            self.input_path_str = str(path)
        else:
            self.input_path_str = path
            path = pathlib.Path(path)
        self.create_dir = create_dir

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
                    f"Directory {self.full} does not exist and create_dir is set to False. Call create() before any other commands!"
                )
        elif not self.full.is_dir():
            raise NotADirectoryError(f"Directory {self.full} is not a directory.")

    @property
    def files(self) -> list["File"]:
        """
        Gets all files in the dir as a list of File objects.
        """
        return [
            File(self.full / file) for file in self.full.iterdir() if file.is_file()
        ]

    @property
    def files_r(self) -> list["File"]:
        """
        Recursively gets all files in the dir, so including files in subdirs, as a list of File objects.
        """
        return [
            File(self.full / file) for file in self.full.rglob("*") if file.is_file()
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
            return [Directory(str(d), False) for d in self.full.iterdir() if d.is_dir()]
        if r:
            return [Directory(str(d), False) for d in self.full.rglob("*") if d.is_dir()]

    def newest_file(self, file_type:list[str]|str|None = None) -> "File":
        """
        Returns the newest file in the dir as a File object.
        Parameters:
            file_type (str): If set, only files with this extension will be returned.
            input the extension incl dot; or a list of them.
        """
        all_files = self.files
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
        all_files = self.files_r
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
                return shutil.copy2(src, dst, follow_symlinks=follow_symlinks)

        if not overwrite:
            shutil.copytree(self.full, target, dirs_exist_ok=True, copy_function=copy_only_new)
        else:
            shutil.copytree(self.full, target, dirs_exist_ok=True)

    def __eq__(self, other) -> bool:
        return self.full == other.full

    def __str__(self):
        return str(self.full)

    def __repr__(self):
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

    def __init__(self, path: str | pathlib.Path):
        self._path_init_str = str(path)

        assert isinstance(path, str) or isinstance(path, pathlib.Path)

        if isinstance(path, pathlib.Path):
            self._path = path
            self._name = path.name
            self._extension = path.suffix
            self._dir = Directory(str(self._path.absolute().parent))
        elif isinstance(path, str):
            if "/" in path:
                self._name = path.rsplit("/", 1)[-1]
                self._dir = Directory(path.rsplit("/", 1)[0], create_dir=True)
            else:
                self._name = path
                self._dir = Directory(os.getcwd())

            self._extension = self._name.split(".")[-1]
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
        return datetime.fromtimestamp(self._path.stat().st_birthtime)

    @property
    def modified(self) -> datetime:
        return datetime.fromtimestamp(self._path.stat().st_mtime)

    def copy(self, new_path: str) -> "File":
        shutil.copy(self._path, new_path)
        return File(new_path)

    def move(self, new_path: str | pathlib.Path) -> "File":
        '''
        Moves the file to the indicated new path.
        If the
        '''
        if isinstance(new_path, str):
            new_path = pathlib.Path(new_path)
        if new_path.exists():
            if '.' in str(new_path):
                new_path = pathlib.Path(str(new_path).split(".")[0]+"_" +str(int(time.time()))+'.'+str(new_path).split(".")[1])
            else:
                new_path = pathlib.Path(str(new_path)+"_" +str(int(time.time())))
        shutil.move(self._path, str(new_path.absolute()))
        return File(new_path)

    def rename(self, new_name: str) -> "File":
        self._path = self._dir.full / new_name
        return File(self._path)

    def delete(self) -> None:
        os.remove(self._path)

    def __eq__(self, other: "File") -> bool:
        return self._path == other.path

    def __str__(self):
        return str(self._path)

    def __repr__(self):
        return str(self._path.absolute())
