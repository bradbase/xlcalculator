
# Excel reference: https://support.office.com/en-us/article/SUMPRODUCT-function-16753e75-9f68-4874-94ac-4d2145a2fd2e

from functools import reduce

from pandas import concat

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

class Sumproduct(KoalaBaseFunction):
    """"""

    @staticmethod
    def sumproduct(range_1, *ranges):
        """"""

        if isinstance(range_1, XLRange):
            range_1 = range_1.value

        range_length = len(range_1.values)

        for range in ranges: # if a range has no values (i.e if it's empty)
            this_range_len = len(range.value.values)
            if range_length != this_range_len:
                raise ExcelError("#VALUE!", "The length of the ranges does not match. Looking for {} and you've given me a range of length {}".format(range_length, this_range_len))

            if this_range_len == 0:
                return 0

        sumproduct_ranges = [range_1]
        for range in ranges:
            for item in range.value.values:
                # If there is an ExcelError inside a Range, sumproduct should output an ExcelError
                if isinstance(item, ExcelError):
                    raise ExcelError("#N/A", "ExcelErrors are present in the sumproduct items")

            sumproduct_ranges.append(range.value)

        sumproduct = concat(sumproduct_ranges, axis=1)

        return sumproduct.prod(axis=1).sum()
