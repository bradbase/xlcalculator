
"""
Representation of a Microsoft Excel cell
"""

import logging
import re
from dataclasses import dataclass, field
from string import ascii_uppercase

from .xltype import XLType
from .formula import XLFormula

def init_list():
    return []


@dataclass
class XLCell(XLType):
    """Representing an Excel Cell"""

    address: str = field(compare=True, hash=True, repr=True)
    _address:  str = field(init=False, repr=False)
    sheet: str = field(init=False, repr=False)
    row: str = field(init=False, repr=False)
    column: str = field(init=False, repr=False)
    column_index: int = field(default=None, init=False, repr=True)
    value: str = field(default=None, repr=True)
    formula: XLFormula = field(default=None, repr=True)
    defined_names: list = field(default_factory=init_list, repr=True) # these are "back-links" to the defined names in Model


    def __post_init__(self):
        """Continue initialisation."""

        self.column_index = XLCell.column_ordinal(self.column)


    @property
    def address(self):
        """Because we have overridden address, we need to return the correct value."""

        return self._address


    @address.setter
    def address(self, address):
        """Overrides the address setter so we can break the address up into consitiuent parts."""
        self._address, self.sheet, self.row, self.column = XLCell.extract_address(address)


    @staticmethod
    def extract_address(address):
        if ":" in address:
            raise Exception("This is a Range {}".format(address))

        address = address.replace('$', '')
        if "!" in address:
            sheet, cell_address = address.split('!')
            column, row = [_f for _f in re.split('([A-Z\$]+)', cell_address) if _f]
            logging.debug("Cell address {} sheet {} row {} column {}".format(address, sheet, row, column))

        return (address, sheet, row, column)


    @staticmethod
    def column_ordinal(column):
        """Calculates the column ordinal from characters."""

        tot = 0
        for i, c in enumerate([c for c in column[::-1] if c != "$"]):
            if c == '$': continue
            tot += (ord(c)-64) * 26 ** i
        return tot


    @staticmethod
    def ordinal_column(num):
        """ """

        if num < 1:
            raise Exception("Number must be larger than 0: %s" % num)

        s = ''
        q = num
        while q > 0:
            (q,r) = divmod(q,26)
            if r == 0:
                q = q - 1
                r = 26
            s = ascii_uppercase[r-1] + s

        return s


    @staticmethod
    def unfix(address):
        """"""

        return address.replace('$', '')


    def __hash__(self):
        """Override the hash builtin to hash the address only."""

        return hash( self._address )


    def __lt__(self, other):
        other_address, other_sheet, other_row, other_column = XLCell.extract_address(other)
        truths = []
        truths.append(self.sheet == other_sheet)
        truths.append(self.row < other_row or XLCell.column_ordinal(self.column) < XLCell.column_ordinal(other_column))

        print("__LT__", truths)

        return all(truths)


    def __le__(self, other):
        other_address, other_sheet, other_row, other_column = XLCell.extract_address(other)
        truths = []
        truths.append(self.sheet == other_sheet)
        truths.append(self.row <= other_row or XLCell.column_ordinal(self.column) <= XLCell.column_ordinal(other_column))

        print("__LE__", truths)

        return all(truths)


    def __gt__(self, other):
        other_address, other_sheet, other_row, other_column = XLCell.extract_address(other)
        truths = []
        truths.append(self.sheet == other_sheet)
        truths.append(self.row > other_row or XLCell.column_ordinal(self.column) > XLCell.column_ordinal(other_column))

        print("__GT__", truths)

        return all(truths)


    def __ge__(self):
        other_address, other_sheet, other_row, other_column = XLCell.extract_address(other)
        truths = []
        truths.append(self.sheet == other_sheet)
        truths.append(self.row >= other_row or XLCell.column_ordinal(self.column) >= XLCell.column_ordinal(other_column))

        print("__GE__", truths)

        return all(truths)


    def __eq__(self, other):
        if self.__class__ == other.__class__:
            truths = []
            truths.append(self.address == other.address)

            return all(truths)
            
        else:
            return False


    def __float__(self):
        return float(self.value)
