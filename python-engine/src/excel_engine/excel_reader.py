"""Small workbook-reading adapter used by catalog and audit workflows."""

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

    # Purpose: Validate the path and select the correct pandas engine.
    def __init__(self, path):
        """Validate the path and select the correct pandas engine."""
        # Compute and store: self.path.
        self.path = self.validate_path(path)
        # Compute and store: self.engine.
        self.engine = self.detect_engine()


    ### Validate the file path
    # Purpose: Validate file existence, type, extension, and read permission.
    # Example output: Path("/tmp/input.xlsx")
    def validate_path(self, path) -> Path:
        """Validate file existence, type, extension, and read permission."""
        # Check this condition before choosing the next branch.
        if path is None:
            # Stop with a clear validation error for invalid input.
            raise ValueError("An Excel file path is required.")

        # Convert the input into a Path object and expand the home directory
        file_path = Path(path).expanduser()

        # Check whether the file exists
        if not file_path.exists():
            # Stop with a clear validation error for invalid input.
            raise FileNotFoundError(
                f"Excel file does not exist: {file_path}"
            )

        # Ensure the path points to a file rather than a directory
        if not file_path.is_file():
            # Stop with a clear validation error for invalid input.
            raise ValueError(
                f"Path is not a regular file: {file_path}"
            )

        # Only support .xls and .xlsx file extensions
        if file_path.suffix.lower() not in ENGINES:
            # Stop with a clear validation error for invalid input.
            raise ValueError(
                "Only .xlsx and .xls files are supported."
            )

        # Check whether the current process has permission to read the file
        if not os.access(file_path, os.R_OK):
            # Stop with a clear validation error for invalid input.
            raise PermissionError(
                f"Excel file is not readable: {file_path}"
            )

        # Return the normalized absolute file path
        return file_path.resolve()

    ###### Dectect the document formatting
    # Purpose: Identify the file format and select the appropriate Excel engine.
    # Example output: "openpyxl"
    def detect_engine(self) -> str:
        """Identify the file format and select the appropriate Excel engine."""

        # Get the file extension in lowercase
        suffix = self.path.suffix.lower()

        # An XLSX workbook is typically a ZIP archive
        if zipfile.is_zipfile(self.path):
            # Check this condition before choosing the next branch.
            if suffix != ".xlsx":
                # Stop with a clear validation error for invalid input.
                raise ValueError(
                    "File format does not match extension."
                )
            # Return the value produced by this step.
            return "openpyxl"

        # Read the file signature to identify the OLE Compound File format
        with self.path.open("rb") as file:
            # Compute and store: signature.
            signature = file.read(len(OLE_SIGNATURE))

        # Select xlrd for a valid OLE-based XLS file
        if signature == OLE_SIGNATURE and suffix == ".xls":
            # Return the value produced by this step.
            return "xlrd"

        # Raise an error if the file format cannot be identified
        raise ValueError(
            f"Unrecognized Excel file format: {self.path.name}"
        )

    ################### read the file
    # Purpose: Read all worksheets and return them as a dictionary of DataFrames.
    # Example output: {"02": DataFrame(...)}
    def read_workbook(self) -> dict[str, pd.DataFrame]:
        """Read all worksheets and return them as a dictionary of DataFrames."""

        # Return the value produced by this step.
        return pd.read_excel(
            self.path,
            sheet_name=None,
            engine=self.engine
        )

    ########## read specific worksheet 
    # Purpose: Read one worksheet with optional header and skipped rows.
    # Example output: DataFrame(rows=61, columns=25)
    def read_sheet(self,sheet_name: str,skiprows=None,header=0) -> pd.DataFrame:
        """Read one worksheet with optional header and skipped rows."""

        # Return the value produced by this step.
        return pd.read_excel(
            self.path,
            sheet_name=sheet_name,
            engine=self.engine,
            skiprows=skiprows,
            header=header
        )


# Purpose: Compatibility wrapper used by the JSON entry point.
# Example output: {"02": DataFrame(...)}
def read_workbook(path: PathLike) -> dict[str, pd.DataFrame]:
    """Compatibility wrapper used by the JSON entry point."""
    # Return the value produced by this step.
    return ExcelReader(path).read_workbook()

