"""Representation of a Microsoft Excel formula
"""
from dataclasses import dataclass, field
from openpyxl.utils.cell import column_index_from_string
from typing import List

from . import ast_nodes, tokenizer, utils
from .tokenizer import ExcelParserTokens


class XLType:
    pass


@dataclass
class XLFormula(XLType):
    """Representing an Excel Formula"""

    formula: str = field(compare=True, hash=True, repr=True)

    sheet_name: str = field(default=None, repr=True)
    reference: str = field(default=None, repr=True)
    evaluate: bool = field(default=True, repr=True)
    tokens: List[tokenizer.f_token] = field(
        init=False, default_factory=list, repr=True)
    terms: List[str] = field(init=False, default_factory=list, repr=True)
    associated_cells: set = field(init=False, default_factory=set, repr=True)
    ast: ast_nodes.ASTNode = field(init=False, default=None)

    def __post_init__(self):
        """Supplimentary initialisation."""
        self.tokens = tokenizer.ExcelParser().getTokens(self.formula).items
        for token in self.tokens:
            if (
                    (token.ttype == ExcelParserTokens.TOK_TYPE_OPERAND)
                    and (token.tsubtype == ExcelParserTokens.TOK_SUBTYPE_RANGE)
                    and (token.tvalue not in self.terms)
            ):
                # Make sure we have a full address.
                term = token.tvalue
                if '!' not in term:
                    term = f'{self.sheet_name}!{term}'
                self.terms.append(term)


@dataclass
class XLCell(XLType):
    """Excel Cell"""

    address: str = field(compare=False, repr=True)

    sheet: str = field(compare=True, hash=True, init=False, repr=False)
    row: str = field(compare=False, hash=False, init=False, repr=False)
    row_index: int = field(compare=True, hash=True, init=False, repr=False)
    column: str = field(compare=False, hash=False, init=False, repr=False)
    column_index: int = field(compare=True, hash=True, init=False, repr=False)
    value: str = field(compare=False, default=None, repr=True)
    formula: XLFormula = field(
        compare=False, default=None, hash=False, repr=True)
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


@dataclass
class XLRange(XLType):
    """Excel Range"""

    address_str: str = field(compare=False, hash=False, repr=True)
    name: str = field(default=None, compare=False, hash=True, repr=True)

    cells: list = field(init=False, compare=True, hash=False, repr=False)
    sheet: str = field(init=False, default="Sheet1", repr=False)
    value: list = field(default=None, repr=True)

    def __post_init__(self):
        if self.name is None:
            self.name = self.address_str
        self.sheet, self.cells = utils.resolve_ranges(self.address_str)

    @property
    def address(self):
        return self.cells
