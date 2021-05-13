from typing import Literal, Optional, Union

from . import xl, func_xltypes
from .func_xltypes import UNUSED, Unused
from .xlerrors import NumExcelError, ValueExcelError

dec = "dec"
Base = Literal[bin, dec, oct, hex]

PERMITTED_DIGITS = {
    bin: set("01"),
    oct: set("01234567"),
    hex: set("0123456789ABCDEFabcdef"),
}


BIT_WIDTHS = {bin: 10, oct: 30, hex: 40}

BASE_NUMBERS = {bin: 2, oct: 8, hex: 16}


BOUNDS = {
    frozenset([bin, oct]): 2 ** 9,
    frozenset([bin, dec]): 2 ** 9,
    frozenset([bin, hex]): 2 ** 9,
    frozenset([oct, dec]): 2 ** 29,
    frozenset([oct, hex]): 2 ** 29,
    frozenset([dec, hex]): 2 ** 39,
}


def handle_places(
        places: Union[Unused, func_xltypes.XlAnything]
) -> Optional[int]:
    if places is UNUSED:
        return None

    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError('The `places` argument cannot be a boolean.')

    places = int(places)
    if not (1 <= places <= 10):
        raise NumExcelError('The number of places must be between 1 and 10.')

    return places


def handle_number(number: func_xltypes.XlAnything, origin) -> Union[int, str]:
    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError('The number cannot be a boolean.')

    if origin == dec:
        return int(number)

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError('Number is not an integer.')

        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number) if number else "0"

    if len(as_str) > 10:
        raise NumExcelError()

    if set(as_str) - PERMITTED_DIGITS[origin]:
        raise NumExcelError('There are invalid characters in the number.')

    return as_str


def pad_zeroes(string, was_negative, places):
    if places is None:
        return string

    desired_length = len(string) if was_negative else places
    if desired_length < len(string):
        raise NumExcelError(
            'The resulting string is longer than the desired length.')

    return string.zfill(desired_length)


def conversion(number, origin, destination, places):
    if origin == dec:
        # if origin base is dec we do not need to convert from it
        value = number
    else:
        # otherwise, convert to int in the appropriate base.
        as_int = int(number, BASE_NUMBERS[origin])

        # magic:
        mask = 1 << BIT_WIDTHS[origin] - 1
        value = (as_int & ~mask) - (as_int & mask)

    # This is another error-check, but it had to be delayed to here rather
    # than in handle_number because we had to convert it from the origin base
    # first.
    bound = BOUNDS[frozenset([origin, destination])]
    if not (-bound <= value < bound):
        raise NumExcelError('The input number is out of bounds')

    # if the destination base is dec, we are done
    if destination == dec:
        return value

    # magic (handles the 2s-complement-like wrapping behavior):
    was_negative = value < 0
    if was_negative:
        value += 1 << BIT_WIDTHS[destination]

    # Otherwise convert to appropriate string representation (return value is
    # (str, bool)).
    result = destination(value)[2:].upper()
    return pad_zeroes(result, was_negative, places)


def convert_bases(number, origin, destination, places=None):
    # If places is None, means it was not passed to this function, i.e. the
    # corresponding Excel function does not take it as an argument (the
    # ___2DEC functions). If it was passed in, we do error checking via the
    # handle_places function.
    if places is not None:
        places = handle_places(places)

    # Next, we do error checking on the number argument. This function does
    # not do any base-conversion logic, it simply raises exceptions if it
    # encounters invalid inputs.  If the origin base is dec, the result is an
    # integer, otherwise it is a string.
    number = handle_number(number, origin)

    # Now we actually convert from one base to another. The result of this can
    # also be either and integer or a string depending on the destination
    # base. If it is a string, it may also be zero-padded depending on the
    # places argument.
    return conversion(number, origin, destination, places)


@xl.register()
@xl.validate_args
def DEC2BIN(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, dec, bin, places)


@xl.register()
@xl.validate_args
def DEC2OCT(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, dec, oct, places)


@xl.register()
@xl.validate_args
def DEC2HEX(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, dec, hex, places)


# the ___2DEC functions give a number not a string, and they do not take a
# `places` parameter
@xl.register()
@xl.validate_args
def BIN2DEC(
        number: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return convert_bases(number, bin, dec)


@xl.register()
@xl.validate_args
def BIN2OCT(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, bin, oct, places)


@xl.register()
@xl.validate_args
def BIN2HEX(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, bin, hex, places)


@xl.register()
@xl.validate_args
def OCT2DEC(
        number: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return convert_bases(number, oct, dec)


@xl.register()
@xl.validate_args
def OCT2BIN(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, oct, bin, places)


@xl.register()
@xl.validate_args
def OCT2HEX(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, oct, hex, places)


@xl.register()
@xl.validate_args
def HEX2DEC(
        number: func_xltypes.XlAnything
) -> func_xltypes.XlNumber:
    return convert_bases(number, hex, dec)


@xl.register()
@xl.validate_args
def HEX2BIN(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, hex, bin, places)


@xl.register()
@xl.validate_args
def HEX2OCT(
        number: func_xltypes.XlAnything,
        places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    return convert_bases(number, hex, oct, places)
