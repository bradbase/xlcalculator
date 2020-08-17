import functools
import inspect
import typing

from . import func_xltypes, xlerrors

COMPATIBILITY = 'EXCEL'
CELL_CHARACTER_LIMIT = 32767

TYPE_TO_CAST = {
    func_xltypes.XlNumber: func_xltypes.Number.cast,
    func_xltypes.XlText: func_xltypes.Text.cast,
    func_xltypes.XlBoolean: func_xltypes.Boolean.cast,
    func_xltypes.XlDateTime: func_xltypes.DateTime.cast,
    func_xltypes.XlArray: func_xltypes.Array.cast,
    func_xltypes.XlExpr: func_xltypes.Expr.cast,
    func_xltypes.XlAnything: func_xltypes.ExcelType.cast_from_native,
}


class Functions(dict):

    def register(self, func, name=None):
        if name is None:
            name = func.__name__
        self[name] = func

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError:
            raise AttributeError(name)


FUNCTIONS = Functions()


def register(name=None):
    """Decorator to register a function."""

    def registerFunction(func):
        FUNCTIONS.register(func, name)
        return func

    return registerFunction


def _validate(vtype, val, name):
    cast = TYPE_TO_CAST.get(vtype, None)
    if cast is not None:
        return cast(val)

    # Support lists with value types
    if getattr(vtype, '__origin__', None) in [list, tuple]:
        itype = vtype.__args__[0]
        if itype != func_xltypes.XlArray:
            val = flatten(val)
        return tuple(filter(
            lambda x: x is not None,
            [_safe_validate(itype, item, name) for item in val]
        ))

    # Support unions
    if getattr(vtype, '__origin__', None) == typing.Union:
        for stype in vtype.__args__:
            try:
                return _validate(stype, val, name)
            except xlerrors.ExcelError:
                pass
        raise xlerrors.ValueExcelError(val)

    return val


def _safe_validate(vtype, val, name):
    try:
        return _validate(vtype, val, name)
    except xlerrors.ExcelError:
        return None


def validate_args(func):

    @functools.wraps(func)
    def validate(*args, **kw):
        sig = inspect.signature(func)
        bound = sig.bind(*args, **kw)
        # 1. Convert all input parameters to Excel Types.
        for pname, value in list(bound.arguments.items()):
            if isinstance(value, xlerrors.ExcelError):
                return value
            try:
                bound.arguments[pname] = _validate(
                    sig.parameters[pname].annotation, value, pname)
            except xlerrors.ExcelError as err:
                return err
        # 2. Run the function to compute the result.
        try:
            res = func(*bound.args, **bound.kwargs)
        except xlerrors.ExcelError as err:
            # Never crash on Excel errors as we want to store them as the cell
            # value.
            return err
        # 3. Convert the result to an Excel type.
        return _validate(sig.return_annotation, res, 'return')

    return validate


def flatten(values):
    """Fully recursive flattening."""
    flat = []
    if isinstance(values, func_xltypes.Array):
        values = values.flat
    for value in values:
        if isinstance(value, func_xltypes.Array):
            flat.extend(value.flat)
        elif isinstance(value, (list, tuple)):
            flat.extend(flatten(value))
        else:
            flat.append(value)
    return flat


def length(values):
    return len(flatten(values))
