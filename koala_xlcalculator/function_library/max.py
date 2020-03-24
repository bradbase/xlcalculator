
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
            max_list = []
            for arg in args:
                if isinstance(arg, XLRange):
                    max_list.append(arg.value.max().max())

                elif isinstance(arg, XLCell):
                    max_list.append(arg.value)

                else:
                    max_list.append(arg.max().max())

            if len(max_list) == 1:
                return max_list[0]

            else:
                return npmaximum(max_list)
