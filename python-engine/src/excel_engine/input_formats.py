"""Names and required fields used by the contractor workbook layout."""

FORMULA_INPUT_FIELDS = (
    "W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K"
)

SHEET_02_COLUMN_MAP = {
    "Unnamed: 0": "section_marker",
    "Unnamed: 1": "item_no",
    "Unnamed: 2": "description",
    "Unnamed: 3": "unit",
    "Unnamed: 5": "W",
    "Unnamed: 6": "H",
    "Unnamed: 7": "W1",
    "Unnamed: 8": "H1",
    "Unnamed: 9": "W2",
    "Unnamed: 10": "H2",
    "Unnamed: 11": "L_M",
    "Unnamed: 12": "R_D",
    "Unnamed: 13": "a",
    "Unnamed: 14": "t",
    "Unnamed: 15": "Q",
    "Unnamed: 16": "K",
}

PRODUCT_INPUT_FORMATS = {
    "rectangular_duct": ("W", "H", "K"),
    "round_duct": ("R_D", "K"),
    "end_cap": ("W", "H", "K"),
    "reducer": ("W", "H", "W1", "H1", "L_M", "K"),
    "square_round_reducer": ("W", "H", "L_M", "R_D", "K"),
    "tee": ("W", "H", "W1", "H1", "W2", "H2", "L_M", "K"),
    "takeoff": ("W", "H", "W1", "H1", "L_M", "K"),
    "elbow": ("W", "H", "L_M", "R_D", "K"),
    "square_elbow": ("W", "H", "R_D", "K"),
    "mfd_valve": ("W", "H", "L_M", "K"),
    "nrd_valve": ("W", "H", "L_M", "K"),
}

__all__ = [
    "FORMULA_INPUT_FIELDS",
    "SHEET_02_COLUMN_MAP",
    "PRODUCT_INPUT_FORMATS",
]
