from enum import Enum


class NameMode(str, Enum):
    FROM_CELL = "from_cell"
    FROM_SHEET_NAME = "from_sheet_name"
    SEQUENTIAL = "sequential"
