
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
    terms: List[xltype.XLType] = field(
        init=False, default_factory=list, repr=True)
    associated_cells: set = field(init=False, default_factory=set, repr=True)
    ast: ast_nodes.ASTNode = field(init=False, default=None)

    def __post_init__(self):
        """Supplimentary initialisation."""
        parser = ExcelParser()

        self.tokens = parser.getTokens(
            self.formula, sheet_name=self.sheet_name).items

        for token in self.tokens:
            if (
                    token.ttype == ExcelParserTokens.TOK_TYPE_OPERAND
                    and token.tsubtype == ExcelParserTokens.TOK_SUBTYPE_RANGE
                    and token.tvalue not in self.terms
            ):
                self.terms.append(token.tvalue)
