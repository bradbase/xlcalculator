
# Excel reference: https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Choose(KoalaBaseFunction):
    """"""

    @staticmethod
    def choose(index_num, *values):
        """"""

        if isinstance(index_num, str):
            index = int(index_num)

        if index <= 0 or index > 254:
            raise ExcelError("#VALUE!", "{} must be between 1 and 254".format( str(index_num) ))

        elif index > len(values):
            raise ExcelError("#VALUE!", "{} must not be larger than the number of values: {}".format( str(index_num), len(values)) )

        else:
            return values[index - 1]
