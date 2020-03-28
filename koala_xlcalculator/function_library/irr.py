
# Excel reference: https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc
# Numpy reference: http://docs.scipy.org/doc/numpy-1.10.0/reference/generated/numpy.irr.html

from numpy_financial import irr as npirr
from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

"""
Guess    Optional. A number that you guess is close to the result of IRR.

Microsoft Excel uses an iterative technique for calculating IRR. Starting with guess, IRR cycles through the calculation until the result is accurate within 0.00001 percent. If IRR can't find a result that works after 20 tries, the #NUM! error value is returned.

In most cases you do not need to provide guess for the IRR calculation. If guess is omitted, it is assumed to be 0.1 (10 percent).

If IRR gives the #NUM! error value, or if the result is not close to what you expected, try again with a different value for guess.
"""


class IRR(KoalaBaseFunction):
    """"""

    @staticmethod
    def irr(values, guess=None):
        """"""
        if isinstance(values, XLRange):
            values = IRR.flatten(values.value.values)

        elif isinstance(values, DataFrame):
            values = IRR.flatten(values.values)

        if guess is not None and guess != 0:
            raise ValueError('guess value for excellib.irr() is %s and not 0' % guess)

        else:
                return npirr(values)
