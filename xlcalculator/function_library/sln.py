
# Excel reference: https://support.office.com/en-us/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8


from .excel_lib import XlCalculatorBaseFunction
from ..xlcalculator_types import XLCell

class SLN(XlCalculatorBaseFunction):
    """"""

    @staticmethod
    def sln(cost, salvage, life):
        """"""

        if isinstance(cost, XLCell):
            cost = cost.value

        if isinstance(salvage, XLCell):
            salvage = salvage.value

        if isinstance(life, XLCell):
            life = life.value

        if isinstance(cost, (int, float)) and isinstance(salvage, (int, float)) and isinstance(life, (int, float)):
            return (cost - salvage) / life
