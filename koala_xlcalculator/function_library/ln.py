
from math import log

from .excel_lib import KoalaBaseFunction

class Ln(KoalaBaseFunction):
    """"""

    def ln(self, a):
        """"""

        if isinstance(a, (list, tuple, np.ndarray) ):
            return [log(x) for x in KoalaBaseFunction.flatten(a)]

        else:
            #print a
            return log(a)
