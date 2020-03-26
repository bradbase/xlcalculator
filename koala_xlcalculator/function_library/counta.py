
import itertools

from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..exceptions import ExcelError

class Counta(KoalaBaseFunction):
    """"""

    @staticmethod
    def counta(arg_1, *args):
        """"""

        def get_count(arg):
            if isinstance(arg, XLRange):
                return len([element for element in [item for item in itertools.chain( *arg.value.values.tolist() ) ] if element not in ['', None] ])

            elif isinstance(arg, DataFrame):
                # I don't like nesting list comprehensions. But here we are...
                return len([element for element in [item for item in itertools.chain( *arg.values.tolist() ) ] if element not in ['', None] ])

            elif arg not in ['', None]:
                return 1

            else:
                return 0


        if arg_1 is None:
            raise ExcelError('#VALUE', 'value1 is required')

        if len(list(args)) > 255:
            raise ExcelError('#VALUE', "Can only have up to 255 supplimentary arguments, you provided {}".format(len(args)))

        total = 0
        total += get_count(arg_1)

        for arg in args:
            total += get_count(arg)

        return total
