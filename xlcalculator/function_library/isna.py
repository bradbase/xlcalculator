
from .excel_lib import XlCalculatorBaseFunction

class Isna(XlCalculatorBaseFunction):
    """"""

    def isna(self, value):
        """"""

        # This function might need more solid testing
        try:

            eval(value)
            return False

        except:
            return True
