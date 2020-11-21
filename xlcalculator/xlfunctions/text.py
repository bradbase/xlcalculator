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
def CONCATENATE(
        *parameters: Tuple[func_xltypes.XlAnything]
) -> func_xltypes.XlText:
    """Use CONCATENATE, one of the text functions, to join two or more
    text strings into one string.

    https://support.microsoft.com/en-us/office/
        concatenate-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d

    IMPORTANT: In Excel 2016, Excel Mobile, and Excel for the web, this
    function has been replaced with the CONCAT function. Although the
    CONCATENATE function is still available for backward compatibility,
    you should consider using CONCAT from now on. This is because CONCATENATE
    may not be available in future versions of Excel.
    """

    return CONCAT(
        [func_xltypes.Text.cast(parameter) for parameter in parameters]
    )


@xl.register()
@xl.validate_args
def FIND(
        find_text: func_xltypes.XlText,
        within_text: func_xltypes.XlText,
        start_num: func_xltypes.Number = 0,
) -> func_xltypes.Number:
    """FIND and FINDB locate one text string within a second text string,
    and return the number of the starting position of the first text string
    from the first character of the second text string.

    https://support.microsoft.com/en-us/office/
        find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628
    """
    index = 1
    within_text_str = str(within_text)
    find_text_str = str(find_text)
    start_num_int = int(start_num)

    # Excel operates in 1-based land, Python is usually 0-based.
    if start_num_int > 0:
        start_num_int = start_num_int - 1

    try:
        index = within_text_str.index(find_text_str, start_num_int) + 1
    except ValueError:
        raise xlerrors.ValueExcelError(
            f"Text {find_text} isn't found in"
            f" {within_text_str[:start_num_int]}")

    return index


@xl.register()
@xl.validate_args
def LEFT(
        text: func_xltypes.XlText,
        num_chars: func_xltypes.XlNumber = 1
) -> func_xltypes.XlText:
    """LEFT returns the first character or characters in a text string,
    based on the number of characters you specify.

    https://support.office.com/en-us/article/
        left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c
    """
    return str(text)[:int(num_chars)]


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
