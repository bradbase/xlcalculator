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
    return xlerrors.NaExcelError()


@xl.register()
def ISNA(cell) -> func_xltypes.Boolean:
    """Returns True if the cell is #N/A.

    Don't call validate_args here because we allow errors to be
    passed in.

    """
    return isinstance(cell, xlerrors.NaExcelError)
