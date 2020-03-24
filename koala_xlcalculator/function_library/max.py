
# Excel reference: https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098

from numpy import maximum as npmaximum

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..koala_types import XLCell


class xMax(KoalaBaseFunction):
    """Finds the maximum of provided values."""

    @staticmethod
    def xmax(*args):
        """Finds the maximum of provided values."""
        # however, if no non numeric cells, return zero (is what excel does)
        if len(args) < 1:
            return 0

        else:
            average_list = []
            for arg in args:
                if isinstance(arg, XLRange):
                    average_list.append(arg.value.max().max())

                elif isinstance(arg, XLCell):
                    average_list.append(arg.value)

                else:
                    average_list.append(arg.max().max())

            return npmaximum(average_list)
