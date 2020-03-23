
# Excel reference: https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89

from numpy import sum as npsum

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..koala_types import XLCell

class xSum(KoalaBaseFunction):
    """"""

    @staticmethod
    def xsum(*args):
        """Ignore non numeric cells and boolean cells."""

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

                else:

                    sum_list.append(arg.sum().sum())

            return sum(sum_list)
