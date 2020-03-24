
# Excel reference: https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6

from numpy import average as npaverage

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..koala_types import XLCell


class Average(KoalaBaseFunction):
    """Find the average (mean) of provided values."""

    @staticmethod
    def average(*args):
        """Find the average (mean) of provided values."""
        # however, if no non numeric cells, return zero (is what excel does)
        if len(args) < 1:
            return 0

        else:
            average_list = []
            for arg in args:
                if isinstance(arg, XLRange):
                    average_list.append(arg.value.mean().mean())

                elif isinstance(arg, XLCell):
                    average_list.append(arg.value)

                else:
                    average_list.append(arg.mean().mean())

            return npaverage(average_list)
