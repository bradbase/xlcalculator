
# https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

class Vlookup(KoalaBaseFunction):
    """"""

    def vlookup(self, lookup_value, table_array, col_index_num, range_lookup=True):
        """"""

        raise Exception("VLOOKUP DOESN'T WORK, XLRANGE IS NOT SUPPORTED")

        if not isinstance(table_array, XLRange):
            return ExcelError('#VALUE', 'table_array should be a Range')

        if col_index_num > table_array.ncols:
            return ExcelError('#VALUE', 'col_index_num is greater than the number of cols in table_array')

        first_column = table_array.get(0, 1)
        result_column = table_array.get(0, col_index_num)

        if not range_lookup:
            if lookup_value not in first_column.values:
                return ExcelError('#N/A', 'lookup_value not in first column of table_array')

            else:
                i = first_column.values.index(lookup_value)
                ref = first_column.order[i]
        else:
            i = None
            for v in first_column.values:
                if lookup_value >= v:
                    i = first_column.values.index(v)
                    ref = first_column.order[i]

                else:
                    break

            if i is None:
                return ExcelError('#N/A', 'lookup_value smaller than all values of table_array')

        return XLRange.find_associated_value(ref, result_column)
