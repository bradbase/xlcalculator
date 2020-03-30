
# Excel reference: https://support.office.com/en-us/article/IFERROR-function-c526fd07-caeb-47b8-8bb6-63f3e417f611


from .excel_lib import KoalaBaseFunction

class Iferror(KoalaBaseFunction):
    """"""

    @staticmethod
    def iferror(value, value_if_error):
        """"""

        if isinstance(value, ExcelError) or value in ErrorCodes:
            return value_if_error

        else:
            return value
