
# https://support.office.com/en-us/article/POWER-function-D3F2908B-56F4-4C3F-895A-07FB519C362A


from numpy import power as nppower

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Power(KoalaBaseFunction):
    """"""

    @staticmethod
    def power(number, power):
        """"""

        return nppower(number, power)
