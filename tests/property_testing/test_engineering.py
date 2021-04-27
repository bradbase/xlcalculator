#%%
from contextlib import contextmanager
import pytest
from hypothesis import settings
from hypothesis.strategies import (
    integers,
    floats,
    one_of,
    none,
    text,
    booleans,
    just,
    sampled_from,
)

from tests.testing import assert_equivalent, Case, parametrize_cases
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
from tests.xlwings_fixtures import fuzz_scalars, given

MAX_EXAMPLES = CONFIG["max-examples"]


def zero_pad_strategy(n: str):
    return sampled_from(list(range(13))).map(lambda x: n.zfill(x))


def integer_variants(n: int):
    # given an integer, return a strategy comprising different representations of that
    # integer: as itself, or as a possibly zero-padded string
    return one_of(just(n), just(str(n)).flatmap(zero_pad_strategy))


def xl_numbers(min_value=None, max_value=None):
    return one_of(
        none(),  # blank cell
        just(""),
        booleans(),
        integers(min_value=min_value, max_value=max_value).flatmap(integer_variants),
        floats(
            min_value=min_value,
            max_value=max_value,
            allow_infinity=False,
            allow_nan=False,
        ).filter(
            lambda x: x > 1e-300
        ),  # guard against tiny floats that excel can't handle
    )


def near(value, width=5):
    return xl_numbers(min_value=value - width, max_value=value + width)


def binary_numbers_as_strings(max_size):
    return text(alphabet=set("01"), min_size=1, max_size=max_size)


def octal_numbers_as_strings(min_size, max_size):
    return text(alphabet=set("01234567"), min_size=min_size, max_size=max_size)


def hex_numbers_as_strings(min_size, max_size):
    return text(
        alphabet=set("0123456789ABCDEFabcdef"), min_size=min_size, max_size=max_size
    )


@pytest.mark.xlwings
@parametrize_cases(
    Case("dec2bin", formula=DEC2BIN, values=given(number=xl_numbers(-550, 550))),
    Case(
        "dec2bin-places",
        formula=DEC2BIN,
        values=given(number=xl_numbers(-550, 550), places=xl_numbers(-2, 12)),
    ),
    Case(
        "dec2oct",
        formula=DEC2OCT,
        values=given(number=one_of(xl_numbers(), near(-(2 ** 29)), near(2 ** 29))),
    ),
    Case(
        "dec2oct-places",
        formula=DEC2OCT,
        values=given(
            number=one_of(xl_numbers(), near(-(2 ** 29)), near(2 ** 29)),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "dec2hex",
        formula=DEC2HEX,
        values=given(number=one_of(xl_numbers(), near(-(2 ** 39)), near(2 ** 39))),
    ),
    Case(
        "dec2hex-places",
        formula=DEC2HEX,
        values=given(
            number=one_of(xl_numbers(), near(-(2 ** 39)), near(2 ** 39)),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "bin2oct",
        formula=BIN2OCT,
        values=given(number=one_of(xl_numbers(), binary_numbers_as_strings(11))),
    ),
    Case(
        "bin2oct-places",
        formula=BIN2OCT,
        values=given(
            number=one_of(xl_numbers(), binary_numbers_as_strings(11)),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "bin2dec",
        formula=BIN2DEC,
        values=given(number=one_of(xl_numbers(), binary_numbers_as_strings(11))),
    ),
    # __2DEC functions do not take a `places` parameter so there is no bin2dec-places
    Case(
        "bin2hex",
        formula=BIN2HEX,
        values=given(number=one_of(xl_numbers(), binary_numbers_as_strings(11))),
    ),
    Case(
        "bin2hex-places",
        formula=BIN2HEX,
        values=given(
            number=one_of(xl_numbers(), binary_numbers_as_strings(11)),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "oct2bin",
        formula=OCT2BIN,
        values=given(
            number=one_of(
                xl_numbers(),
                octal_numbers_as_strings(1, 4),
                octal_numbers_as_strings(9, 11),
            )
        ),
    ),
    Case(
        "oct2bin-places",
        formula=OCT2BIN,
        values=given(
            number=one_of(
                xl_numbers(),
                octal_numbers_as_strings(1, 4),
                octal_numbers_as_strings(9, 11),
            ),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "oct2dec",
        formula=OCT2DEC,
        values=given(number=one_of(xl_numbers(), octal_numbers_as_strings(1, 11))),
    ),
    Case(
        "oct2hex",
        formula=OCT2HEX,
        values=given(number=one_of(xl_numbers(), octal_numbers_as_strings(1, 11))),
    ),
    Case(
        "oct2hex-places",
        formula=OCT2HEX,
        values=given(
            number=one_of(xl_numbers(), octal_numbers_as_strings(1, 11)),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "hex2bin",
        formula=HEX2BIN,
        values=given(
            number=one_of(
                xl_numbers(),
                hex_numbers_as_strings(1, 3),
                hex_numbers_as_strings(9, 11),
            )
        ),
    ),
    Case(
        "hex2bin-places",
        formula=HEX2BIN,
        values=given(
            number=one_of(
                xl_numbers(),
                hex_numbers_as_strings(1, 3),
                hex_numbers_as_strings(9, 11),
            ),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "hex2oct",
        formula=HEX2OCT,
        values=given(number=one_of(xl_numbers(), hex_numbers_as_strings(1, 11))),
    ),
    Case(
        "hex2oct-places",
        formula=HEX2OCT,
        values=given(
            number=one_of(xl_numbers(), hex_numbers_as_strings(1, 11)),
            places=xl_numbers(-2, 12),
        ),
    ),
    Case(
        "hex2dec",
        formula=HEX2DEC,
        values=given(number=one_of(xl_numbers(), hex_numbers_as_strings(1, 11))),
    ),
)
def test_against_xlwings(excel_workbook, formula, values):
    fuzz_scalars(excel_workbook, formula, values)


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
