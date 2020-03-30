
# Excel reference: https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c
# Excel reference: https://support.office.com/en-us/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7
# Excel reference: https://support.office.com/en-us/article/rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53

from decimal import Decimal, getcontext, ROUND_HALF_UP, ROUND_UP, ROUND_DOWN

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class xRound(KoalaBaseFunction):
    """"""

    @staticmethod
    def roundup(number, num_digits=0):
        """Round down"""

        return xRound.xround(number, num_digits=num_digits, rounding=ROUND_UP)


    @staticmethod
    def rounddown(number, num_digits=0):
        """Rounding up"""

        return xRound.xround(number, num_digits=num_digits, rounding=ROUND_DOWN)


    @staticmethod
    def xround(number, num_digits=0, rounding=ROUND_HALF_UP):
        """Rounding half up"""

        if not xRound.is_number(number):
            return ExcelError("#VALUE!", "{} is not a number".format(str(number)))

        if not xRound.is_number(num_digits):
            return ExcelError("#VALUE!", "{} is not a number".format(str(num_digits)))

        number = Decimal(str(number))
        dc = getcontext()
        dc.rounding = rounding
        ans = round(number, num_digits)

        return float( ans )
