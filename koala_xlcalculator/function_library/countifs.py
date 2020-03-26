
# Excel reference: https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLRange

class Countifs(KoalaBaseFunction):
    """"""

    def countifs(self, *args):
        """"""

        raise Exception("COUNTIFS DOESN'T WORK, XLRANGE IS NOT SUPPORTED")

        arg_list = list(args)
        l = len(arg_list)

        if l % 2 != 0:
            raise ExcelError('#VALUE!', 'excellib.countifs() must have a pair number of arguments, here %d' % l)


        if l >= 2:
            indexes = KoalaBaseFunction.find_corresponding_index(args[0].values, args[1]) # find indexes that match first layer of countif

            remaining_ranges = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 0] # get only ranges
            remaining_criteria = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 1] # get only criteria

            # verif that all Ranges are associated COULDNT MAKE THIS WORK CORRECTLY BECAUSE OF RECURSION
            # association_type = None

            # temp = [args[0]] + remaining_ranges

            # for index, range in enumerate(temp): # THIS IS SHIT, but works ok
            #     if type(range) == Range and index < len(temp) - 1:
            #         asso_type = range.is_associated(temp[index + 1])

            #         print 'asso', asso_type
            #         if association_type is None:
            #             association_type = asso_type
            #         elif associated_type != asso_type:
            #             association_type = None
            #             break

            # print 'ASSO', association_type

            # if association_type is None:
            #     return ValueError('All items must be Ranges and associated')

            filtered_remaining_ranges = []

            for range in remaining_ranges: # filter items in remaining_ranges that match valid indexes from first countif layer
                filtered_remaining_cells = []
                filtered_remaining_range = []

                for index, item in enumerate(range.values):
                    if index in indexes:
                        filtered_remaining_cells.append(range.addresses[index]) # reconstructing cells from indexes
                        filtered_remaining_range.append(item) # reconstructing values from indexes

                # WARNING HERE
                filtered_remaining_ranges.append(XLRange(filtered_remaining_cells, filtered_remaining_range))

            new_tuple = ()

            for index, range in enumerate(filtered_remaining_ranges): # rebuild the tuple that will be the argument of next layer
                new_tuple += (range, remaining_criteria[index])

            return min(self.countifs(*new_tuple), len(indexes)) # only consider the minimum number across all layer responses

        else:
            return float('inf')
