
from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLCell

COMPATIBILITY = 'EXCEL'

class Right(KoalaBaseFunction):
    """"""


    @staticmethod
    def right(text, number_of_chars = 1):
        """"""

        if COMPATIBILITY == "ANTHILL":
            return Right.anthill(text, number_of_chars)
        else:
            return Right.excel(text, number_of_chars)


    @staticmethod
    def excel(text, number_of_chars):

        if not isinstance(number_of_chars, int):
            raise ExcelError("#VALUE!", "RIGHT function number of cars must be int, you've given me {}".format( type(number_of_chars) ))

        if isinstance(text, XLCell):
            text = str(text.value)
        else:
            text = str(text)

        return text[-number_of_chars:]


    @staticmethod
    def anthill(text, number_of_chars):

        #TODO: hack to deal with naca section numbers
        if isinstance(text) or isinstance(text, str):
            return text[-number_of_chars:]

        else:
            # TODO: get rid of the decimal
            return str(int(text))[-number_of_chars:]
