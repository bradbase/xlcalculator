
# https://support.office.com/en-ie/article/sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf

from numpy import sqrt

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLCell

class Sqrt(KoalaBaseFunction):
    """"""

    @staticmethod
    def sqrt(number):
        """"""

        if isinstance(number, XLCell):
            number = number.value

        if isinstance(number, (int, float)):
            if number < 0:
                return ExcelError('#NUM!', '{} must be non-negative'.format( number ))

            return sqrt(number)
            
        else:
            return ExcelError('#VALUE!', '{} must be a number. You gave me a {}'.format( number, type(number) ))
