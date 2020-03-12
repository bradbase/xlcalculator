
# Excel reference: https://support.office.com/en-us/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class SLN(KoalaBaseFunction):
    """"""

    def sln(self, cost, salvage, life):
        """"""

        for arg in [cost, salvage, life]:
            if isinstance(arg, ExcelError) or arg in KoalaBaseFunction.ErrorCodes:
                return arg

        return (cost - salvage) / life
