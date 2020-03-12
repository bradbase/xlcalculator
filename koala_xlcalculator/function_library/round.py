
# Excel reference: https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c
# Excel reference: https://support.office.com/en-us/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7


from decimal import Decimal, ROUND_HALF_UP, ROUND_UP

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Round(KoalaBaseFunction):
    """"""

    def roundup(self, number, num_digits=0):
        """"""

        return self.xround(number, num_digits=num_digits, rounding=ROUND_UP)


    def xround(self, number, num_digits=0, rounding=ROUND_HALF_UP):
        """"""

        if not KoalaBaseFunction.is_number(number):
            return ExcelError("#VALUE!", "{} is not a number".format(str(number)))

        if not KoalaBaseFunction.is_number(num_digits):
            return ExcelError("#VALUE!", "{} is not a number".format(str(num_digits)))

        number = float(number) # if you don't Spreadsheet.dump/load, you might end up with Long numbers, which Decimal doesn't accept

        if num_digits >= 0: # round to the right side of the point
            return float(Decimal(repr(number)).quantize(Decimal(repr(pow(10, -num_digits))), rounding=rounding))
            # see https://docs.python.org/2/library/functions.html#round
            # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5

        else:
            return round(number, num_digits)
