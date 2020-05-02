
# Excel reference: https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89

from numpy import sum as npsum

from .excel_lib import XlCalculatorBaseFunction
from ..xlcalculator_types import XLRange
from ..xlcalculator_types import XLCell

class xSum(XlCalculatorBaseFunction):
    """Sum all provided values."""

    @staticmethod
    def xsum(*args):
        """Sum all provided values."""

        # however, if no non numeric cells, return zero (is what excel does)
        if len(args) < 1:
            return 0

        else:
            sum_list = []
            for arg in args:
                if isinstance(arg, XLRange):
                    sum_list.append(arg.value.sum().sum())

                elif isinstance(arg, XLCell):
                    sum_list.append(arg.value)

                elif isinstance(arg, (int, float)):
                    sum_list.append(arg)

                else:
                    sum_list.append(arg.sum().sum())

            return sum(sum_list)
