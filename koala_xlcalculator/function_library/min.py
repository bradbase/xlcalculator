
# Excel reference: https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152


from .excel_lib import KoalaBaseFunction

class xMin(KoalaBaseFunction):
    """"""

    def xmin(self, *args):
        """"""
        # Ignore non numeric cells and boolean cells.

        values = KoalaBaseFunction.extract_numeric_values(*args)

        # however, if no non numeric cells, return zero (is what excel does)
        if len(values) < 1:
            return 0

        else:
            return min(values)
