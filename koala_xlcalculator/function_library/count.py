
# Excel reference: https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c


from .excel_lib import KoalaBaseFunction
from ..koala_types import XLRange

class Count(KoalaBaseFunction):
    """"""

    def count(self, *args):
        """"""

        raise Exception("COUNT DOESN'T WORK, XLRANGE DOESN'T SUPPORT .values")

        l = list(args)
        total = 0

        for arg in l:
            if isinstance(arg, XLRange):
                total += len([x for x in arg.values if KoalaBaseFunction.is_number(x) and type(x) is not bool]) # count inside a list

            elif KoalaBaseFunction.is_number(arg): # int() is used for text representation of numbers
                total += 1

        return total
