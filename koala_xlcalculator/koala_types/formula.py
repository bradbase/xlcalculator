
"""
Representation of a Microsoft Excel formula
"""

import logging
from dataclasses import dataclass, field
import re
from typing import List

from ..read_excel import f_token
from ..read_excel import ExcelParser


def init_ranges():
    """Default factory to initialise Formula.ranges."""

    return []


def init_tokens():
    """Default factory to initialise Formula.tokens."""

    return []


@dataclass
class XLFormula():
    """Representing an Excel Formula"""

    formula: str = field(compare=True, hash=True, repr=True)
    sheet_name: str = field(default=None, repr=True)
    return_type: str = field(default='value', repr=True)
    reference: str = field(default=None, repr=True)
    evaluate: bool = field(default=True, repr=True)
    tokens: List[f_token] = field(init=False, default_factory=init_tokens, repr=True)
    ranges: List[str] = field(init=False, default_factory=init_ranges, repr=True)
    python_code: str = field(init=False, default=None, repr=True)


    def __post_init__(self):
        """Supplimentary initialisation."""

        if self.return_type not in ['array']:
            parser = ExcelParser()
            self.tokens = parser.getTokens(self.formula, sheet_name=self.sheet_name).items

            for item in self.tokens:
                if item.ttype == 'operand' and item.tsubtype == 'range':
                    self.ranges.append(item.tvalue)


    def __hash__(self):
        """Override the hash builtin to hash the formula only."""

        return hash( (self.formula, self.return_type, self.reference, self.evaluate, self.python_code) )


    def __eq__(self, other):
        truths = []
        truths.append(self.__class__ == other.__class__)
        truths.append(self.formula == other.formula)
        truths.append(self.return_type == other.return_type)
        truths.append(self.reference == other.reference)
        truths.append(self.evaluate == other.evaluate)
        truths.append(self.python_code == other.python_code)

        return all(truths)
