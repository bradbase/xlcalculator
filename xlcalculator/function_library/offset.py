
# Excel reference: https://support.office.com/en-us/article/OFFSET-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66


from .excel_lib import XlCalculatorBaseFunction
from ..exceptions import ExcelError

class Offset(XlCalculatorBaseFunction):
    """"""

    def offset(self, reference, rows, cols, height=None, width=None):
        """This function accepts a list of addresses
        Maybe think of passing a Range as first argument.
        """

        raise Exception("OFFSET DOESN'T WORK, XLRANGE ISN'T SUPPORTED")

        for i in [reference, rows, cols, height, width]:
            if isinstance(i, ExcelError) or i in ErrorCodes:
                return i

        rows = int(rows)
        cols = int(cols)

        # get first cell address of reference
        if is_range(reference):
            ref = resolve_range(reference, should_flatten = True)[0][0]

        else:
            ref = reference

        ref_sheet = ''
        end_address = ''

        if '!' in ref:
            ref_sheet = ref.split('!')[0] + '!'
            ref_cell = ref.split('!')[1]

        else:
            ref_cell = ref

        found = re.search(CELL_REF_RE, ref)
        new_col = col2num(found.group(1)) + cols
        new_row = int(found.group(2)) + rows

        if new_row <= 0 or new_col <= 0:
            return ExcelError('#VALUE!', 'Offset is out of bounds')

        start_address = str(num2col(new_col)) + str(new_row)

        if (height is not None and width is not None):
            if type(height) != int:
                return ExcelError('#VALUE!', '%d must not be integer' % height)

            if type(width) != int:
                return ExcelError('#VALUE!', '%d must not be integer' % width)

            if height > 0:
                end_row = new_row + height - 1

            else:
                return ExcelError('#VALUE!', '%d must be strictly positive' % height)

            if width > 0:
                end_col = new_col + width - 1

            else:
                return ExcelError('#VALUE!', '%d must be strictly positive' % width)

            end_address = ':' + str(num2col(end_col)) + str(end_row)

        elif height and not width or not height and width:
            return ExcelError('Height and width must be passed together')

        return ref_sheet + start_address + end_address
