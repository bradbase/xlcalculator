
# Excel reference: https://support.office.com/en-us/article/CONCATENATE-function-8F8AE884-2CA8-4F7A-B093-75D702BEA31D
# Important: In Excel 2016, Excel Mobile, and Excel Online, this function has
# been replaced with the CONCAT function. Although the CONCATENATE function is
# still available for backward compatibility, you should consider using CONCAT
# from now on. This is because CONCATENATE may not be available in future
# versions of Excel.
#
# BE AWARE; there are functional differences between CONACTENATE AND CONCAT
#


from .excel_lib import XlCalculatorBaseFunction
from ..exceptions import ExcelError

class Concatenate(XlCalculatorBaseFunction):
    """"""

    def concatenate(self, *args):
        """"""

        if tuple(XlCalculatorBaseFunction.flatten(args)) != args:
            return ExcelError('#VALUE', 'Could not process arguments %s' % (args))

        cat_string = ''.join(str(a) for a in args)

        if len(cat_string) > XlCalculatorBaseFunction.CELL_CHARACTER_LIMIT:
            return ExcelError('#VALUE', 'Too long. concatentaed string should be no longer than %s but is %s' % (XlCalculatorBaseFunction.CELL_CHARACTER_LIMIT, len(cat_String)))

        return cat_string
