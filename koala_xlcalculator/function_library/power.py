
# https://support.office.com/en-us/article/POWER-function-D3F2908B-56F4-4C3F-895A-07FB519C362A


import numpy as np

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Power(KoalaBaseFunction):
    """"""

    def power(self, number, power):
        """"""

        if number == power == 0:
            # Really excel?  What were you thinking?
            return ExcelError('#NUM!', 'Number and power cannot both be zero' % str(number))

        if power < 1 and number < 0:
            return ExcelError('#NUM!', '%s must be non-negative' % str(number))

        return np.power(number, power)
