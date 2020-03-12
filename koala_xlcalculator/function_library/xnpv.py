
# Excel reference: https://support.office.com/en-us/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7


import numpy as np

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class XNPV(KoalaBaseFunction):
    """"""

    def xnpv(self, *args):
        """"""

        rate = args[0]
        # ignore non numeric cells and boolean cells
        values = KoalaBaseFunction.extract_numeric_values(args[1])
        dates = KoalaBaseFunction.extract_numeric_values(args[2])
        if len(values) != len(dates):
            return ExcelError("#NUM!", '`values` range must be the same length as `dates` range in XNPV, %s != %s' % (len(values), len(dates)))

        xnpv = 0
        for v, d in zip(values, dates):
            xnpv += v / np.power(1.0 + rate, (d - dates[0]) / 365)

        return xnpv
