
# Excel reference: https://support.office.com/en-us/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568


from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange

class NPV(KoalaBaseFunction):
    """"""

    def npv(self, *args):
        """"""

        raise Exception("NPV DOESN'T WORK, XLRANGE DOESN'T SUPPORT .values")

        discount_rate = args[0]
        cashflow = args[1]

        if isinstance(cashflow, XLRange):
            cashflow = cashflow.values

        return sum([float(x)*(1+discount_rate)**-(i+1) for (i, x) in enumerate(cashflow)])
