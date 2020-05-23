"""Representation of a Microsoft Excel cell
"""
import re
from dataclasses import dataclass, field
from openpyxl.utils.cell import SHEET_TITLE, COORD_RE, column_index_from_string

from . import xltype, formula as xlformula, utils


@dataclass
class XLCell(xltype.XLType):
    """Representing an Excel Cell"""

    address: str = field(compare=False, repr=True)

    sheet: str = field(compare=True, hash=True, init=False, repr=False)
    row: str = field(compare=False, hash=False, init=False, repr=False)
    row_index: int = field(compare=True, hash=True, init=False, repr=False)
    column: str = field(compare=False, hash=False, init=False, repr=False)
    column_index: int = field(compare=True, hash=True, init=False, repr=False)
    value: str = field(compare=False, default=None, repr=True)
    formula: xlformula.XLFormula = field(compare=False, default=None, hash=False, repr=True)
    # These are "back-links" to the defined names in Model.
    defined_names: list = field(compare=False, default_factory=list, repr=True)

    def __post_init__(self):
        self.sheet, self.column, self.row = utils.resolve_address(self.address)
        self.column_index = column_index_from_string(self.column)
        self.row_index = int(self.row)

    def __float__(self):
        return float(self.value)

    def __hash__(self):
        # XXX: Also should not be needed.
        return hash((self.sheet, self.row_index, self.column_index))

