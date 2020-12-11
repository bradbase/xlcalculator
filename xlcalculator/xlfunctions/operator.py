from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def OP_MUL(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return left * right


@xl.register()
@xl.validate_args
def OP_DIV(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    if right == 0:
        raise xlerrors.DivZeroExcelError()
    return left / right


@xl.register()
@xl.validate_args
def OP_ADD(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return left + right


@xl.register()
@xl.validate_args
def OP_SUB(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return left - right


@xl.register()
def OP_EQ(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    return left == right


@xl.register()
def OP_NE(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    return left != right


@xl.register()
@xl.validate_args
def OP_GT(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left > right


@xl.register()
@xl.validate_args
def OP_LT(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left < right


@xl.register()
@xl.validate_args
def OP_GE(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left >= right


@xl.register()
@xl.validate_args
def OP_LE(
        left: func_xltypes.XlAnything,
        right: func_xltypes.XlAnything
) -> func_xltypes.XlBoolean:
    if isinstance(left, func_xltypes.Blank) or \
            isinstance(right, func_xltypes.Blank):
        return False
    return left <= right


@xl.register()
@xl.validate_args
def OP_NEG(
        right: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    return -1 * right


@xl.register()
@xl.validate_args
def OP_PERCENT(
        left: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    return left * 0.01
