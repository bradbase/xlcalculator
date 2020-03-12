
from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Counta(KoalaBaseFunction):
    """"""

    def counta(self, range):
        """"""

        if isinstance(range, ExcelError) or range in ErrorCodes:
            if range.value == "#NULL":
                return 0

            else:
                return range # return the Excel Error
                # raise Exception('ExcelError other than #NULL passed to excellib.counta()')

        else:
            return len([x for x in range.values if x != None])
