
from .excel_lib import KoalaBaseFunction

class Isna(KoalaBaseFunction):
    """"""

    def isna(self, value):
        """"""

        # This function might need more solid testing
        try:

            eval(value)
            return False

        except:
            return True
