
# Excel reference: https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6


from .excel_lib import KoalaBaseFunction

class Average(KoalaBaseFunction):
    """"""

    def average(self, *args):
        """"""
        # Ignore non numeric cells and boolean cells.

        values = KoalaBaseFunction.extract_numeric_values(*args)

        return sum(values) / len(values)
