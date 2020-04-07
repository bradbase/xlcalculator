
"""
Representation of a Microsoft Excel formula
"""

import logging
from dataclasses import dataclass, field
import re
from typing import List
from string import ascii_uppercase
from copy import copy

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
    shared_formula: str = field(default=None, compare=False, repr=False)
    shared_formula_offset: int = field(default=None, compare=False, repr=False)
    shared_formula_range: str = field(default=None, compare=False, repr=False)


    def __post_init__(self):
        """Supplimentary initialisation."""

        if self.return_type not in ['array', 'shared']:
            parser = ExcelParser()
            self.tokens = parser.getTokens(self.formula, sheet_name=self.sheet_name).items

            for item in self.tokens:
                if item.ttype == 'operand' and item.tsubtype == 'range':
                    self.ranges.append(item.tvalue)

        elif self.return_type == 'shared':
            parser = ExcelParser()
            self.tokens = parser.getTokens(self.formula, sheet_name=self.sheet_name).items

            for item in self.tokens:
                if item.ttype == 'operand' and item.tsubtype == 'range':
                    direction = XLFormula.formula_direction(self.shared_formula_range)
                    sheet, cell_address = item.tvalue.split('!')
                    orig_cell = copy(cell_address)
                    column, row = XLFormula.split_col_row(cell_address)
                    row = int(row)

                    if direction == 'rows':
                        row += self.shared_formula_offset
                    else:
                        column_ordinal = XLFormula.column_ordinal(column)
                        column = XLFormula.ordinal_column(column_ordinal + self.shared_formula_offset)

                    item.tvalue = "{}!{}{}".format(sheet, column, row)
                    self.ranges.append(item.tvalue)
                    self.formula = self.formula.replace(orig_cell, "{}{}".format(column, row))


    @staticmethod
    def split_col_row(address):
        return  [_f for _f in re.split('([A-Z\$]+)', address) if _f]


    @staticmethod
    def formula_direction(address):
        """"""

        cells = []
        address = address.replace(' ', '')
        start_cell_address, end_cell_address = address.split(':')

        start_column, start_row = XLFormula.split_col_row(start_cell_address)
        end_column, end_row = XLFormula.split_col_row(end_cell_address)
        start_row = int(start_row)
        end_row = int(end_row)
        start_column_ordinal = XLFormula.column_ordinal(start_column)
        end_column_ordinal = XLFormula.column_ordinal(end_column)

        # only one column involved, we can concentrate on the rows.

        if start_column_ordinal == end_column_ordinal:
            return "rows" # increase rows, down

        # Multiple columns involved
        else:
            return "columns" # increase columns, across


    @staticmethod
    def column_ordinal(column):
        """Calculates the column ordinal from characters."""

        tot = 0
        for i, c in enumerate([c for c in column[::-1] if c != "$"]):
            if c == '$': continue
            tot += (ord(c) - 64) * 26 ** i
        return tot


    @staticmethod
    def ordinal_column(num):
        """ """

        if num < 1:
            raise Exception("Number must be larger than 0: {}".format(num))

        s = ''
        q = num
        while q > 0:
            (q, r) = divmod(q, 26)
            if r == 0:
                q = q - 1
                r = 26
            s = ascii_uppercase[r - 1] + s

        return s


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
