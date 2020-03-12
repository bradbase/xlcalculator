
# Excel reference: https://support.office.com/en-us/article/SUMPRODUCT-function-16753e75-9f68-4874-94ac-4d2145a2fd2e


from functools import reduce

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

class Sumproduct(KoalaBaseFunction):
    """"""

    def sumproduct(self, *ranges):
        """"""

        raise Exception("SUMPRODUCT DOESN'T WORK, XLRANGE ISN'T SUPPORTED")

        range_list = list(ranges)

        for r in range_list: # if a range has no values (i.e if it's empty)
            if len(r.values) == 0:
                return 0

        for range in range_list:
            for item in range.values:
                # If there is an ExcelError inside a Range, sumproduct should output an ExcelError
                if isinstance(item, ExcelError):
                    return ExcelError("#N/A", "ExcelErrors are present in the sumproduct items")

        reduce(check_length, range_list) # check that all ranges have the same size

        return reduce(lambda X, Y: X + Y, reduce(lambda x, y: XLRange.apply_all('multiply', x, y), range_list).values)
