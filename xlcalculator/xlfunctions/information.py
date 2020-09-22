from typing import Tuple

from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def ISBLANK(cell: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Returns TRUE if the cell is empty.
    """
    return isinstance(cell, func_xltypes.Blank) or cell.value == ''


@xl.register()
@xl.validate_args
def ISTEXT(cell: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Returns True if the cell is text.
    """
    return isinstance(cell, func_xltypes.Text)


@xl.register()
@xl.validate_args
def NA() -> xlerrors.ExcelError:
    return xlerrors.ERROR_CODE_NA


@xl.register()
@xl.validate_args
def ISNA(cell: func_xltypes.XlAnything) -> func_xltypes.Boolean:
    """Returns True if the cell is #N/A.

    This can be fooled by text of the type #N/A
    """
    return cell.value == xlerrors.ERROR_CODE_NA
