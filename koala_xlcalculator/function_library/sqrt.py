
# https://support.office.com/en-ie/article/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf


import numpy as np

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Sqrt(KoalaBaseFunction):
    """"""

    def sqrt(self, number):
        """"""

        if number < 0:
            raise ExcelError('#NUM!', '%s must be non-negative' % str(index_num))

        return np.sqrt(number)
