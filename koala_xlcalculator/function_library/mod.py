
# Excel Reference: https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Mod(KoalaBaseFunction):
    """"""

    @staticmethod
    def mod(nb, q):
        """"""

        if not isinstance(nb, int):
            return ExcelError("#VALUE!", "{} is not an integer".format(str(nb)))

        elif not isinstance(q, int):
            return ExcelError("#VALUE!", "{} is not an integer".format(str(q)))

        else:
            return nb % q
