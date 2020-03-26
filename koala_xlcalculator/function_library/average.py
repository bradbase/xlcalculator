
# Excel reference: https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6

import logging
import itertools

from numpy import average as npaverage
from pandas import DataFrame

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

                    average_list.extend([item for item in itertools.chain( *arg.value.values.tolist() ) ])

                elif isinstance(arg, XLCell):
                    average_list.append(arg.value)

                elif isinstance(arg, (int, float)):
                    average_list.append(arg)

                elif isinstance(arg, DataFrame):
                    average_list.extend([item for item in itertools.chain( *arg.values.tolist() ) ])

            logging.debug("AVERAGE: {}".format(average_list))

            return npaverage(average_list)
