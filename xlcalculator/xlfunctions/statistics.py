from typing import Tuple

from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def AVERAGE(
        *numbers: Tuple[func_xltypes.Number]
) -> func_xltypes.Number:
    """Returns the average (arithmetic mean) of the arguments.

    https://support.office.com/en-us/article/
        average-function-047bac88-d466-426c-a32b-8f33eb960cf6
    """
    numbers = xl.flatten(numbers)

    # If no non numeric cells, return zero (is what excel does)
    if len(numbers) < 1:
        return 0

    return sum(numbers) / len(numbers)


@xl.register()
@xl.validate_args
def COUNT(*values) -> func_xltypes.Number:
    """Counts the number of cells that contain numbers, and counts numbers
    within the list of arguments.

    https://support.office.com/en-us/article/
        count-function-a59cd7fc-b623-4d93-87a4-d23bf411294c
    """
    values = xl.flatten(values)
    if not len(values) or values[0] is None:
        raise xlerrors.ValueExcelError('value1 is required')

    if len(values) > 255:
        raise xlerrors.ValueExcelError(
            f"Can only have up to 255 supplimentary arguments. "
            f"Provided: {len(values)}")

    return len(list(filter(func_xltypes.Number.is_type, values)))


@xl.register()
@xl.validate_args
def COUNTA(*values):
    """Counts the number of cells that are not empty in a range.

    https://support.office.com/en-us/article/
        counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509
    """
    values = xl.flatten(values)
    if not len(values) or values[0] is None:
        raise xlerrors.NullExcelError('value1 is required')

    if len(values) > 256:
        raise xlerrors.ValueExcelError(
            f"Can only have up to 256 arguments. "
            f"Provided: {len(values)}")

    cells = list(filter(lambda x: not func_xltypes.Blank.is_blank(x), values))
    return len(cells)


@xl.register()
@xl.validate_args
def MAX(*numbers: Tuple[func_xltypes.Number]):
    """Returns the largest value in a set of values.

    https://support.office.com/en-us/article/
        max-function-e0012414-9ac8-4b34-9a47-73e662c08098
    """
    # If no non numeric cells, return zero (is what excel does)
    if len(numbers) < 1:
        return 0

    return max(filter(func_xltypes.Number.is_type, numbers))


@xl.register()
@xl.validate_args
def MIN(*numbers: Tuple[func_xltypes.Number]):
    """Returns the smallest number in a set of values.

    https://support.office.com/en-us/article/
        min-function-61635d12-920f-4ce2-a70f-96f202dcc152
    """
    # If no non numeric cells, return zero (is what excel does)
    if len(numbers) < 1:
        return 0

    return min(filter(func_xltypes.Number.is_type, numbers))
