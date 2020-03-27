
# Excel reference: https://support.office.com/en-us/article/counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509

import itertools

from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..exceptions import ExcelError

class Counta(KoalaBaseFunction):
    """"""

    @staticmethod
    def counta(arg):
        """The COUNTA function counts the number of cells that are not empty in a range."""

        if arg is None:
            raise ExcelError('#VALUE', 'value1 is required')

        if not isinstance(arg, (XLRange, DataFrame)):
            raise ExcelError('#VALUE', "COUNTA only takes a range, you provided {}".format(type(args)))


        if isinstance(arg, XLRange):
            return len([element for element in [item for item in itertools.chain( *arg.value.values.tolist() ) ] if element not in ['', None] ])

        elif isinstance(arg, DataFrame):
            # I don't like nesting list comprehensions. But here we are...
            return len([element for element in [item for item in itertools.chain( *arg.values.tolist() ) ] if element not in ['', None] ])
