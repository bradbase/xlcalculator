
"""
Representation of a Microsoft Excel range
"""

import logging
import re
from dataclasses import dataclass, field

from .cell import XLCell
from .formula import XLFormula


@dataclass
class XLRange():
    """"""

    name: str = field(compare=True, hash=True, repr=True)
    cells: list = field(compare=True, hash=True, repr=True)
    _cells:  list = field(init=False, repr=False)
    sheet: str = field(init=False, default=None, repr=False)
    row: str = field(init=False, default=None, repr=False)
    value: list = field(default=None, repr=True)
    formula: XLFormula = field(default=None, repr=True)
    length: int = field(init=False, default=0, repr=False)


    @property
    def address(self):
        """Because we have overridden address, we need to return the correct value."""

        return self._cells


    @property
    def cells(self):
        """Because we have overridden address, we need to return the correct value."""

        return self._cells


    @cells.setter
    def cells(self, cells):
        """"""
        # we store cell addresses in a range in (essentially) the same format as they are defined.
        # eg; a 2D list. A1, A2, A3, C1, C2, C3 is [[A1, A2, A3],[C1, C2, C3]]

        range_cells = []
        cells = XLCell.unfix(cells)

        # multiple cell groups in this range (there are gaps eg; Sheet1!A1:A5,C1:C5,E1:E5)
        if ',' in cells:
            addresses = cells.split(',')
            for address in addresses:
                range_cells.append( XLRange.cell_address_infill(address)[0] )

        # only one cell group in this range (no gaps eg; Sheet1!A1:E5)
        else:
            range_cells = XLRange.cell_address_infill(cells)

        self._cells = range_cells
        self.length = len(range_cells[0])


    @staticmethod
    def cell_address_infill(address):
        """"""

        cells = []
        start_address, end_cell_address = address.split(':')
        sheet, start_cell_address = start_address.split('!')
        end_address = "{}!{}".format(sheet, end_cell_address)

        start_column, start_row = [_f for _f in re.split('([A-Z\$]+)', start_cell_address) if _f]
        end_column, end_row = [_f for _f in re.split('([A-Z\$]+)', end_cell_address) if _f]
        start_row = int(start_row)
        end_row = int(end_row)
        start_column_ordinal = XLCell.column_ordinal(start_column)
        end_column_ordinal = XLCell.column_ordinal(end_column)

        # only one column involved, we can concentrate on the rows.

        if start_column_ordinal == end_column_ordinal:
            row = []
            for row_ordinal in range(start_row, end_row + 1):
                row.append("{}!{}{}".format(sheet, XLCell.ordinal_column(start_column_ordinal), row_ordinal))
            cells.append(row)

        # Multiple columns involved
        else:
            for column_ordinal in range(start_column_ordinal, end_column_ordinal + 1):
                row = []
                for row_ordinal in range(start_row, end_row + 1):
                    row.append("{}!{}{}".format(sheet, XLCell.ordinal_column(column_ordinal), row_ordinal))
                cells.append(row)

        return cells


    def __hash__(self):
        """Override the hash builtin to hash the address only."""

        return hash( self.name, self.sheet, self.row, self.formula, self.length )


    def __eq__(self, other):
        truths = []
        truths.append(self.__class__ == other.__class__)
        truths.append( self.name == other.name )
        truths.append( self.sheet == other.sheet )
        truths.append( self.row == other.row )
        truths.append( self.formula == other.formula )
        truths.append( self.length == other.length )

        return all(truths)
