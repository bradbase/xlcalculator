
"""
Representation of a Microsoft Excel formula
"""
from dataclasses import dataclass, field
import re
from typing import Any, List

from . import xltype, ast_nodes
from ..read_excel import f_token
from ..read_excel import ExcelParser
from ..read_excel import ExcelParserTokens


@dataclass
class XLFormula(xltype.XLType):
    """Representing an Excel Formula"""

    formula: str = field(compare=True, hash=True, repr=True)

    sheet_name: str = field(default=None, repr=True)
    reference: str = field(default=None, repr=True)
    evaluate: bool = field(default=True, repr=True)
    tokens: List[f_token] = field(init=False, default_factory=list, repr=True)
    terms: List[str] = field(init=False, default_factory=list, repr=True)
    associated_cells: set = field(init=False, default_factory=set, repr=True)
    ast: ast_nodes.ASTNode = field(init=False, default=None)

    def __post_init__(self):
        """Supplimentary initialisation."""
        self.tokens = ExcelParser().getTokens(self.formula).items
        for token in self.tokens:
            if (
                    token.ttype == ExcelParserTokens.TOK_TYPE_OPERAND
                    and token.tsubtype == ExcelParserTokens.TOK_SUBTYPE_RANGE
                    and token.tvalue not in self.terms
            ):
                # Make sure we have a full address.
                term = token.tvalue
                if '!' not in term:
                    term = f'{self.sheet_name}!{term}'
                self.terms.append(term)
