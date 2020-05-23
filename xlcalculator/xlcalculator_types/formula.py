
"""
Representation of a Microsoft Excel formula
"""
from dataclasses import dataclass, field
import re
from typing import List

from .xltype import XLType
from ..read_excel import f_token
from ..read_excel import ExcelParser
from ..read_excel import ExcelParserTokens


@dataclass
class XLFormula(XLType):
    """Representing an Excel Formula"""

    formula: str = field(compare=True, hash=True, repr=True)

    sheet_name: str = field(default=None, repr=True)
    reference: str = field(default=None, repr=True)
    evaluate: bool = field(default=True, repr=True)
    tokens: List[f_token] = field(init=False, default_factory=list, repr=True)
    terms: List[XLType] = field(init=False, default_factory=list, repr=True)
    python_code: str = field(init=False, default=None, repr=True)
    associated_cells: set = field(init=False, default_factory=set, repr=True)

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
