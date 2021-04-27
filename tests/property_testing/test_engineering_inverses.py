from hypothesis import given, settings
from hypothesis.strategies import integers

from xlcalculator.xlfunctions.engineering import (
    DEC2BIN,
    DEC2OCT,
    DEC2HEX,
    BIN2OCT,
    BIN2DEC,
    BIN2HEX,
    OCT2BIN,
    OCT2DEC,
    OCT2HEX,
    HEX2BIN,
    HEX2OCT,
    HEX2DEC,
)

from tests.conftest import CONFIG


MAX_EXAMPLES = CONFIG["max-examples"]


def to_bin(x):
    return bin(x)[2:]


def to_oct(x):
    return oct(x)[2:]


def to_hex(x):
    return hex(x)[2:]


def ten_chars_or_fewer(x):
    return len(x) <= 10


@given(binary_string=integers(min_value=0).map(to_bin).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_bin2oct_and_oct2bin_are_inverses(binary_string):
    # Note that in Python, bin(1023) == "0b1111111111", but in Excel, DEC2BIN(1023) gives
    # a #NUM! error while BIN2DEC(1111111111) == -1. So the numbers in the @given
    # decorator are not the same as how Excel interprets the input. This test just says
    # that OCT2BIN and BIN2OCT are inverses of each other, it says nothing about what
    # actual values they produce or what they mean. The same is true of the other inverse
    # property tests later on in this file.
    assert binary_string == OCT2BIN(BIN2OCT(binary_string))


@given(binary_string=integers(min_value=0).map(to_bin).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_bin2hex_and_hex2bin_are_inverses(binary_string):
    assert binary_string == HEX2BIN(BIN2HEX(binary_string))


@given(binary_string=integers(min_value=0).map(to_bin).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_bin2dec_and_dec2bin_are_inverses(binary_string):
    assert binary_string == DEC2BIN(BIN2DEC(binary_string))


@given(octal_string=integers(min_value=0).map(to_oct).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_oct2dec_and_dec2oct_are_inverses(octal_string):
    assert octal_string == DEC2OCT(OCT2DEC(octal_string))


@given(octal_string=integers(min_value=0).map(to_oct).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_oct2hex_and_hex2oct_are_inverses(octal_string):
    assert octal_string == HEX2OCT(OCT2HEX(octal_string))


@given(decimal_integer=integers(min_value=-(2 ** 39), max_value=2 ** 39 - 1))
@settings(max_examples=MAX_EXAMPLES)
def test_hex2dec_and_dec2hex_are_inverses(decimal_integer):
    assert decimal_integer == HEX2DEC(DEC2HEX(decimal_integer))
