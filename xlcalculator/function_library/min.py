
# Excel reference: https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152

from numpy import minimum as npminimum

from .excel_lib import XlCalculatorBaseFunction
from ..xlcalculator_types import XLRange
from ..xlcalculator_types import XLCell


class xMin(XlCalculatorBaseFunction):
    """Finds the minimum of provided values."""

    @staticmethod
    def xmin(*args):
        """Finds the minimum of provided values."""
        # however, if no non numeric cells, return zero (is what excel does)
        if len(args) < 1:
            return 0

        else:
            min_list = []
            for arg in args:
                if isinstance(arg, XLRange):
                    min_list.append(arg.value.min().min())

                elif isinstance(arg, XLCell):
                    min_list.append(arg.value)

                elif isinstance(arg, (int, float)):
                    min_list.append(arg)
                else:
                    min_list.append(arg.min().min())

            if len(min_list) == 1:
                return min_list[0]

            else:
                return npminimum(*min_list)
