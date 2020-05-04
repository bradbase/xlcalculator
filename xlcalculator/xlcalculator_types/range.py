
"""
Representation of a Microsoft Excel range
"""

import logging
import re
from dataclasses import dataclass, field

from .xltype import XLType
from .cell import XLCell
from .formula import XLFormula


@dataclass
class XLRange(XLType):
    """"""

    name: str = field(compare=True, hash=True, repr=True)
    cells: list = field(compare=True, hash=True, repr=True)
    _cells:  list = field(init=False, repr=False)
    sheet: str = field(init=False, default="Sheet1", repr=False)
    value: list = field(default=None, repr=True)



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
        sheet = None

        # multiple cell groups in this range (there are gaps eg; Sheet1!A1:A5,C1:C5,E1:E5)
        if ',' in cells:
            sheet_count = cells.count('!')
            if sheet_count > 0:
                sheet = cells.split('!')[0]

            addresses = cells.split(',')
            counter = 0
            for address in addresses:
                xlrange = XLRange.cell_address_infill(address, sheet=self.sheet)
                if range_cells == []:
                    range_cells.extend(xlrange)
                else:
                    for index in range(0, len(range_cells)):
                        range_cells[index].extend(xlrange[index])

        # only one cell group in this range (no gaps eg; Sheet1!A1:E5)
        else:
            range_cells = XLRange.cell_address_infill(cells, sheet=self.sheet)

        self._cells = range_cells


    @staticmethod
    def cell_address_infill(address, sheet=None):
        """"""

        cells = []
        address = address.replace(' ', '')

        start_address, end_cell_address = address.split(':')
        if start_address.count("!") > 0:
            sheet, start_cell_address = start_address.split('!')
        else:
            start_cell_address = start_address
            if sheet is None:
                message = "I need a sheet name for this range. {}".format(address)
                logging.error(message)
                raise Exception(message)

        if sheet is not None:
            end_address = "{}!{}".format(sheet, end_cell_address)
        else:
            end_address = "{}".format(end_cell_address)

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
                if sheet is not None:
                    row.append(["{}!{}{}".format(sheet, XLCell.ordinal_column(start_column_ordinal), row_ordinal)])
                else:
                    row.append(["{}{}".format(XLCell.ordinal_column(start_column_ordinal), row_ordinal)])

            cells.extend(row)

        # Multiple columns involved
        else:
            for row_ordinal in range(start_row, end_row + 1):
                column = []
                for column_ordinal in range(start_column_ordinal, end_column_ordinal + 1):
                    if sheet is not None:
                        column.append("{}!{}{}".format(sheet, XLCell.ordinal_column(column_ordinal), row_ordinal))
                    else:
                        column.append("{}{}".format(XLCell.ordinal_column(column_ordinal), row_ordinal))
                cells.append(column)

        return cells


    def __hash__(self):
        """Override the hash builtin to hash the address only."""

        return hash( self.name )


    def __eq__(self, other):
        truths = []
        truths.append(self.__class__ == other.__class__)
        truths.append( self.name == other.name )

        return all(truths)
