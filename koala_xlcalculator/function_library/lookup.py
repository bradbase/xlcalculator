
# Excel reference: https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb


from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError

class Lookup(KoalaBaseFunction):
    """"""

    def lookup(self, value, lookup_range, result_range=None):
        """"""

        if not isinstance(value, (int, float)):
            return Exception("Non numeric lookups ({}) not yet supported".format(value))

        # TODO: note, may return the last equal value

        # index of the last numeric value
        lastnum = -1
        for i,v in enumerate(lookup_range.values):
            if isinstance(v,(int,float)):
                if v > value:
                    break

                else:
                    lastnum = i

        output_range = result_range.values if result_range is not None else lookup_range.values

        if lastnum < 0:
            raise ExcelError('#VALUE!', 'No numeric data found in the lookup range')

        else:
            if i == 0:
                raise ExcelError('#VALUE!', 'All values in the lookup range are bigger than %s' % value)

            else:
                if i >= len(lookup_range)-1:
                    # return the biggest number smaller than value
                    return output_range[lastnum]

                else:
                    return output_range[i-1]
