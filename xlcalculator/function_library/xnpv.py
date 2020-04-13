
# Excel reference: https://support.office.com/en-us/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7

from pandas import DataFrame
from numpy import power
from numpy import sum as npsum

from .excel_lib import XlCalculatorBaseFunction
from ..exceptions import ExcelError
from ..xlcalculator_types import XLRange
from ..xlcalculator_types import XLCell


class XNPV(XlCalculatorBaseFunction):
    """"""

    @staticmethod
    def xnpv(rate, values, dates):
        """"""

        if isinstance(rate, XLCell):
            rate = (float(rate.value))

        else:
            rate = float(rate)

        if isinstance(values, XLRange):
            values = XNPV.flatten( values.value.values )

        elif isinstance(values, DataFrame):
            values = XNPV.flatten( values.values )
        else:
            return ExcelError("#VALUE!", '`values` must be a DataFrame or XLRange, you gave me {}'.format(type(values)))

        if isinstance(dates, XLRange):
            dates = XNPV.flatten( dates.value.values )

        elif isinstance(dates, DataFrame):
            dates = XNPV.flatten( dates.values )
        else:
            return ExcelError("#VALUE!", '`dates` must be a DataFrame or XLRange, you gave me {}'.format(type(dates)))

        # ignore non numeric cells and boolean cells
        if len(values) != len(dates):
            return ExcelError("#NUM!", '`values` range must be the same length as `dates` range in XNPV, %s != %s'.format(len(values), len(dates)))

        def npv(value, date):
            return value / power(1.0 + rate, (date - dates[0]) / 365)

        return npsum([npv(value, date) for value, date in zip(values, dates)])
