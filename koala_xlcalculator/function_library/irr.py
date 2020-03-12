
# Excel reference: https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc
# Numpy reference: http://docs.scipy.org/doc/numpy-1.10.0/reference/generated/numpy.irr.html

import numpy as np

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

class IRR(KoalaBaseFunction):
    """"""

    def irr(self, values, guess=None):
        """"""

        raise Exception("IRR DOESN'T WORK, XLRANGE IS NOT SUPPORTED")

        if (isinstance(values, XLRange)):
            values = values.values

        if guess is not None and guess != 0:
            raise ValueError('guess value for excellib.irr() is %s and not 0' % guess)

        else:
            try:

                return np.irr(values)

            except Exception as e:
                return ExcelError('#NUM!', e)
