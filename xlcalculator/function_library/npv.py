
# Excel reference: https://support.office.com/en-us/article/npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568

from pandas import DataFrame
from numpy_financial import npv as npnpv

from .excel_lib import XlCalculatorBaseFunction
from ..xlcalculator_types import XLRange
from ..xlcalculator_types import XLCell

class NPV(XlCalculatorBaseFunction):
    """"""

    @staticmethod
    def npv(discount_rate, *args):
        """"""

        if len(args) < 1:
            raise Exception("NPV needs a value_1")

        cashflow = []

        if isinstance(discount_rate, XLCell):
            discount_rate = discount_rate.value

        for item in args:
            if isinstance(item, XLRange):
                cashflow.extend( NPV.flatten( item.value ) )

            elif isinstance(item, XLCell):
                cashflow.append( item.value )

            elif isinstance(item, DataFrame):
                cashflow.extend( NPV.flatten( item.values ) )

            elif isinstance(item, (int, float)):
                cashflow.append(item)


        if NPV.COMPATIBILITY == 'PYTHON':
            return npnpv(discount_rate, cashflow)

        else:
            return sum([float(x)*(1+discount_rate)**-(i+1) for (i, x) in enumerate(cashflow)])
