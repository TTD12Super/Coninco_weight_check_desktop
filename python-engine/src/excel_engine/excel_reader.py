from pathlib import Path
from typing import Union
import os
import zipfile
import pandas as pd

# Allow the input path to be either a string or a Path object
PathLike = Union[str, Path]

# Map Excel file extensions to their corresponding reading engines
ENGINES = {
    ".xlsx": "openpyxl",
    ".xls": "xlrd",
}

# Signature used to identify an OLE Compound File (.xls)
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class ExcelReader:
    """Read Excel workbooks in .xls and .xlsx formats."""

    def __init__(self, path):
        self.path = self.validate_path(path)
        self.engine = self.detect_engine()


    ### Validate the file path
    def validate_path(self, path) -> Path:
        if path is None:
            raise ValueError("An Excel file path is required.")

        # Convert the input into a Path object and expand the home directory
        file_path = Path(path).expanduser()

        # Check whether the file exists
        if not file_path.exists():
            raise FileNotFoundError(
                f"Excel file does not exist: {file_path}"
            )

        # Ensure the path points to a file rather than a directory
        if not file_path.is_file():
            raise ValueError(
                f"Path is not a regular file: {file_path}"
            )

        # Only support .xls and .xlsx file extensions
        if file_path.suffix.lower() not in ENGINES:
            raise ValueError(
                "Only .xlsx and .xls files are supported."
            )

        # Check whether the current process has permission to read the file
        if not os.access(file_path, os.R_OK):
            raise PermissionError(
                f"Excel file is not readable: {file_path}"
            )

        # Return the normalized absolute file path
        return file_path.resolve()

    ###### Dectect the document formatting
    def detect_engine(self) -> str:
        """Identify the file format and select the appropriate Excel engine."""

        # Get the file extension in lowercase
        suffix = self.path.suffix.lower()

        # An XLSX workbook is typically a ZIP archive
        if zipfile.is_zipfile(self.path):
            if suffix != ".xlsx":
                raise ValueError(
                    "File format does not match extension."
                )
            return "openpyxl"

        # Read the file signature to identify the OLE Compound File format
        with self.path.open("rb") as file:
            signature = file.read(len(OLE_SIGNATURE))

        # Select xlrd for a valid OLE-based XLS file
        if signature == OLE_SIGNATURE and suffix == ".xls":
            return "xlrd"

        # Raise an error if the file format cannot be identified
        raise ValueError(
            f"Unrecognized Excel file format: {self.path.name}"
        )

    ################### read the file
    def read_workbook(self) -> dict[str, pd.DataFrame]:
        """Read all worksheets and return them as a dictionary of DataFrames."""

        return pd.read_excel(
            self.path,
            sheet_name=None,
            engine=self.engine
        )

    ########## read specific worksheet 
    def read_sheet(self,sheet_name: str,skiprows=None,header=0) -> pd.DataFrame:


        return pd.read_excel(
            self.path,
            sheet_name=sheet_name,
            engine=self.engine,
            skiprows=skiprows,
            header=header
        )


def read_workbook(path: PathLike) -> dict[str, pd.DataFrame]:
    """Compatibility wrapper used by the JSON entry point."""
    return ExcelReader(path).read_workbook()

