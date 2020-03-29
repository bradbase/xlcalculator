
# https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

from pandas import DataFrame

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange
from ..koala_types import XLCell

class VLookup(KoalaBaseFunction):
    """"""

    @staticmethod
    def vlookup(lookup_value, table_array, col_index_num, range_lookup=False):
        """"""

        if range_lookup:
            raise Exception("Excact match only supported at the moment.")

        if isinstance(lookup_value, XLCell):
            lookup_value = lookup_value.value

        if isinstance( table_array, XLRange ):
            table_array = table_array.value

        if col_index_num > len(table_array):
            raise ExcelError('#VALUE', 'col_index_num is greater than the number of cols in table_array')

        table_array = table_array.set_index(0)

        if not range_lookup:
            if lookup_value not in table_array.index:
                raise ExcelError('#N/A', 'lookup_value not in first column of table_array')

            else:
                return table_array.loc[lookup_value].values[0]

        # else:
        #     i = None
        #     for v in first_column.values:
        #         if lookup_value >= v:
        #             i = first_column.values.index(v)
        #             ref = first_column.order[i]
        #
        #         else:
        #             break
        #
        #     if i is None:
        #         raise ExcelError('#N/A', 'lookup_value smaller than all values of table_array')
        #
        # return XLRange.find_associated_value(ref, result_column)
