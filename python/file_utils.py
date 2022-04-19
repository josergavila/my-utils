from pathlib import Path
from typing import Optional, Union


def check_if_file_exists(filename: Union[str, Path]) -> bool:
    """check if a file exists"""
    filename = Path(filename)
    return filename.is_file()
