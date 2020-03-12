
# Excel reference: https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34


from .excel_lib import KoalaBaseFunction

class Countif(KoalaBaseFunction):
    """"""

    def countif(self, range, criteria):
        """"""

        # WARNING:
        # - wildcards not supported
        # - support of strings with >, <, <=, =>, <> not provided

        valid = KoalaBaseFunction.find_corresponding_index(range.values, criteria)

        return len(valid)
