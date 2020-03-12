
from .excel_lib import KoalaBaseFunction

class Right(KoalaBaseFunction):
    """"""

    def right(self, text, n):
        """"""

        #TODO: hack to deal with naca section numbers
        if isinstance(text) or isinstance(text, str):
            return text[-n:]

        else:
            # TODO: get rid of the decimal
            return str(int(text))[-n:]
