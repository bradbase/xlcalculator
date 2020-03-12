
from .excel_lib import KoalaBaseFunction

class Istext(KoalaBaseFunction):
    """"""

    def istext(self, value):
        """"""

        return type(value) == str
