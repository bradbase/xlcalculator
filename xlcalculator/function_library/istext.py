
from .excel_lib import XlCalculatorBaseFunction

class Istext(XlCalculatorBaseFunction):
    """"""

    def istext(self, value):
        """"""

        return type(value) == str
