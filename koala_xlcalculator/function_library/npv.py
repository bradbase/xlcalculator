
# Excel reference: https://support.office.com/en-us/article/npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568

from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..koala_types import XLCell

class NPV(KoalaBaseFunction):
    """"""

    @staticmethod
    def npv(rate, *args):
        """"""

        discount_rate = rate
        if len(args) < 1:
            raise Exception("NPV needs a value_1")

        cashflow = []

        for item in args:
            if isinstance(item, XLRange):
                cashflow.extend( NPV.flatten( item.value ) )

            elif isinstance(item, XLCell):
                cashflow.append( item.value )

            elif isinstance(item, DataFrame):
                cashflow.extend( NPV.flatten( item.values ) )

            elif isinstance(item, (int, float)):
                cashflow.append(item)


        return sum([float(x)*(1+discount_rate)**-(i+1) for (i, x) in enumerate(cashflow)])
