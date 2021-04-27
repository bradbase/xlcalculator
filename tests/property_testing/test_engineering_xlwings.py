import pytest
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

from tests.testing import Case, parametrize_cases
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

try:
    import xlwings
except ImportError:
    pytestmark = pytest.mark.skip("xlwings is not installed; skipping module")
    from hypothesis import given
else:
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
