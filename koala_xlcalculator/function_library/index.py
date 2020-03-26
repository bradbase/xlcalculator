
from .excel_lib import KoalaBaseFunction

class Index(KoalaBaseFunction):
    """"""

    def index(self, my_range, row, col=None): # Excel reference: https://support.office.com/en-us/article/INDEX-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
        """"""

        for i in [my_range, row, col]:
            if isinstance(i, ExcelError) or i in ErrorCodes:
                return i

        row = int(row) if row is not None else row
        col = int(col) if col is not None else col

        if isinstance(my_range, Range):
            cells = my_range.addresses
            nr = my_range.nrows
            nc = my_range.ncols

        else:
            cells, nr, nc = my_range
            if nr > 1 or nc > 1:
                a = np.array(cells)
                cells = a.flatten().tolist()

        nr = int(nr)
        nc = int(nc)

        if type(cells) != list:
            raise ExcelError("#VALUE!", "{} must be a list".format(str(cells)))

        if row is not None and not is_number(row):
            raise ExcelError("#VALUE!", "{} must be a number".format(str(row)))

        if row == 0 and col == 0:
            raise ExcelError("#VALUE!", "No index asked for Range")

        if row is not None and row > nr:
            raise ExcelError("#VALUE!", "Index {} out of range".format(row) )

        if nr == 1:
            col = row if col is None else col
            return cells[int(col) - 1]

        if nc == 1:
            return cells[int(row) - 1]

        else: # could be optimised
            if col is None or row is None:
                raise ExcelError("#VALUE!", "Range is 2 dimensional, can not reach value with 1 arg as None")

            if not is_number(col):
                raise ExcelError("#VALUE!", "{} must be a number".format(str(col)))

            if col > nc:
                raise ExcelError("#VALUE!", "Index {} out of range".format(col))

            indices = list(range(len(cells)))

            if row == 0: # get column
                filtered_indices = [x for x in indices if x % nc == col - 1]
                filtered_cells = [cells[i] for i in filtered_indices]

                return filtered_cells

            elif col == 0: # get row
                filtered_indices = [x for x in indices if int(x / nc) == row - 1]
                filtered_cells = [cells[i] for i in filtered_indices]

                return filtered_cells

            else:
                return cells[(row - 1)* nc + (col - 1)]
