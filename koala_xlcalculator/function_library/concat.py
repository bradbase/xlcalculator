
# https://support.office.com/en-us/article/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2

import itertools

from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange


class Concat(KoalaBaseFunction):
    """"""

    @staticmethod
    def concat(*args):
        """"""
        if len(args) > 254:
            raise ExcelError("#VALUE!", "Can't concat more than 254 arguments. You provided {}".format( len(args) ))

        joinable = []
        for item in args:
            if isinstance(item, str):
                joinable.append(item)

            elif isinstance(item, DataFrame):
                joinable.extend( [str(item) for item in itertools.chain( *item.values.tolist() ) ] )

            elif isinstance(item, XLRange):
                joinable.extend( [str(item) for item in itertools.chain( *item.value.values.tolist() ) ] )

            elif isinstance(item, list):
                joinable.append( list )

        return ''.join(joinable)
