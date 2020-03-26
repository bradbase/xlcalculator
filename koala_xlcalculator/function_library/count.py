
# Excel reference: https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

import itertools

from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange
from ..exceptions import ExcelError

class Count(KoalaBaseFunction):
    """"""

    @staticmethod
    def count(arg_1, *args):
        """"""

        def get_count(arg):
            if isinstance(arg, XLRange):
                return len([x for x in [str(item) for item in itertools.chain( *arg.value.values.tolist() ) ] if KoalaBaseFunction.is_number(x) and type(x) is not bool ])

            elif isinstance(arg, DataFrame):
                # I don't like nesting list comprehensions. But here we are...
                return len([x for x in [str(item) for item in itertools.chain( *arg.values.tolist() ) ] if KoalaBaseFunction.is_number(x) and type(x) is not bool ])

            elif KoalaBaseFunction.is_number(arg): # int() is used for text representation of numbers
                return 1


        if arg_1 is None:
            raise ExcelError('#VALUE', 'value1 is required')

        if len(list(args)) > 255:
            raise ExcelError('#VALUE', "Can only have up to 255 supplimentary arguments, you provided {}".format(len(args)))

        total = 0
        total += get_count(arg_1)

        for arg in args:
            total += get_count(arg)

        return total
