from typing import Tuple

from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def CONCAT(
        *texts: Tuple[func_xltypes.XlText]
) -> func_xltypes.XlText:
    """The CONCAT function combines the text from multiple ranges and/or
    strings, but it doesn't provide delimiter or IgnoreEmpty arguments.

    https://support.office.com/en-us/article/
        concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2
    """
    if len(texts) > 254:
        raise xlerrors.ValueExcelError(
            f"Can't concat more than 254 arguments. Provided: {len(texts)}")

    texts = xl.flatten(texts)
    return ''.join([
        str(text) for text in xl.flatten(texts)
    ])


@xl.register()
@xl.validate_args
def MID(
        text: func_xltypes.XlText,
        start_num: func_xltypes.Number,
        num_chars: func_xltypes.Number
) -> func_xltypes.XlText:
    """Returns a specific number of characters from a text string, starting
    at the position you specify, based on the number of characters you specify.

    https://support.office.com/en-us/article/
        mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028
    """
    text = str(text)

    if len(text) > xl.CELL_CHARACTER_LIMIT:
        raise xlerrors.ValueExcelError(
            f'Text is too long. Is {len(text)} but needs to '
            f'be {xl.CELL_CHARACTER_LIMIT} or less.')

    start_num = int(start_num)
    if start_num < 1:
        raise xlerrors.NumExcelError(f'{start_num} is < 1')

    num_chars = int(num_chars)
    if num_chars < 0:
        raise xlerrors.NumExcelError(f'{num_chars} is < 0')

    start_idx = start_num - 1
    return text[start_idx:start_idx + num_chars]


@xl.register()
@xl.validate_args
def RIGHT(
        text: func_xltypes.XlText,
        num_chars: func_xltypes.XlNumber = 1
) -> func_xltypes.XlText:
    """Returns the last character or characters in a text string.

    https://support.office.com/en-us/article/
        right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f
    """
    return str(text)[-int(num_chars):]
