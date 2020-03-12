
# Excel reference: https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098


from .excel_lib import KoalaBaseFunction

class xMax(KoalaBaseFunction):
    """"""

    def xmax(self, *args):
        """"""
        # Ignore non numeric cells and boolean cells.

        values = KoalaBaseFunction.extract_numeric_values(*args)

        # however, if no non numeric cells, return zero (is what excel does)
        if len(values) < 1:
            return 0

        else:
            return max(values)
