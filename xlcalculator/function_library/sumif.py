
# Excel reference: https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b


from .excel_lib import XlCalculatorBaseFunction

class SumIf(XlCalculatorBaseFunction):
    """"""

    def sumif(self, range, criteria, sum_range=None):
        """"""
        # WARNING:
        # - wildcards not supported
        # - doesn't really follow 2nd remark about sum_range length

        if not isinstance(range, Range):
            return TypeError("{} must be a Range".format( str(range) ) )

        if isinstance(criteria, Range) and not isinstance(criteria , (str, bool)): # ugly...
            return 0

        indexes = XlCalculatorBaseFunction.find_corresponding_index(range.values, criteria)

        if sum_range:
            if not isinstance(sum_range, Range):
                return TypeError("{} must be a Range".format(str(sum_range)))

            def f(x):
                return sum_range.values[x] if x < sum_range.length else 0

            return sum(map(f, indexes))

        else:
            return sum([range.values[x] for x in indexes])
