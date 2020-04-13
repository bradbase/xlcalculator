
# Excel reference: https://support.office.com/en-us/article/ln-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f

from math import log

from numpy import ndarray

from ..xlcalculator_types import XLCell
from .excel_lib import XlCalculatorBaseFunction

class Ln(XlCalculatorBaseFunction):
    """"""

    @staticmethod
    def ln(val):
        """"""

        if isinstance(val, XLCell ):
            return log(val.value)

        elif isinstance(val, (int, float)):
            return log(val)

        else:
            return ExcelError("#VALUE!", "{} must be an XLCell, int or float. You've given me {}".format( val, type(val) ))
