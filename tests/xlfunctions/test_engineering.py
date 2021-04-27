import pytest
from xlcalculator.xlfunctions import engineering
from xlcalculator.xlfunctions.xlerrors import NumExcelError, ValueExcelError

from ..testing import Case, parametrize_cases, assert_equivalent


all_funcs = [
    engineering.BIN2OCT,
    engineering.BIN2DEC,
    engineering.BIN2HEX,
    engineering.OCT2BIN,
    engineering.OCT2DEC,
    engineering.OCT2HEX,
    engineering.DEC2BIN,
    engineering.DEC2OCT,
    engineering.DEC2HEX,
    engineering.HEX2BIN,
    engineering.HEX2OCT,
    engineering.HEX2DEC,
]

all_dec2_funcs = [engineering.DEC2BIN, engineering.DEC2OCT, engineering.DEC2HEX]

all_non_dec2_funcs = [x for x in all_funcs if x not in all_dec2_funcs]

all_2dec_funcs = [engineering.BIN2DEC, engineering.OCT2DEC, engineering.HEX2DEC]

all_non_2dec_funcs = [x for x in all_funcs if x not in all_2dec_funcs]
all_funcs_taking_places = all_non_2dec_funcs


def strings_are_equal(x, y):
    return str(x) == str(y)


@parametrize_cases(
    Case(number=None, expected=0),
    Case(number=0, expected=0),
    Case(number=0.0, expected=0),
    Case(number="0", expected=0),
    Case(number=1, expected=1),
    Case(number="1", expected=1),
    Case(number=True, expected=ValueExcelError),
    Case(number=False, expected=ValueExcelError),
)
@parametrize_cases(*[Case(func=i) for i in all_funcs])
def test_all_funcs(func, number, expected):
    assert_equivalent(func(number), expected, normalize=strings_are_equal)


@parametrize_cases(
    Case(number="", expected=0),
    Case(number="0.0", expected=NumExcelError),
    Case(number="00000000001", expected=NumExcelError),
    Case(number=1.5, expected=NumExcelError),
    Case(number=-1, expected=NumExcelError),
    Case(number="nonsense", expected=NumExcelError),
)
@parametrize_cases(*[Case(func=i) for i in all_non_dec2_funcs])
def test_all_non_dec2_funcs(func, number, expected):
    assert_equivalent(func(number), expected, normalize=strings_are_equal)


@parametrize_cases(
    Case(number="", expected=ValueExcelError),
    Case(number="0.0", expected=0),
    Case(number="00000000001", expected=1),
    Case(number=1.5, expected=1),
    Case(number="nonsense", expected=ValueExcelError),
    Case(number=1.1, expected=1),
    Case(number=1.9, expected=1),
)
@parametrize_cases(*[Case(func=i) for i in all_dec2_funcs])
def test_all_dec2_funcs(func, number, expected):
    assert_equivalent(func(number), expected, normalize=strings_are_equal)


@parametrize_cases(
    Case(number=None, places=None, expected=NumExcelError),
    Case(number=None, places=0, expected=NumExcelError),
    Case(number=0, places=None, expected=NumExcelError),
    Case(number=0, places=0, expected=NumExcelError),
    Case(number=1, places=None, expected=NumExcelError),
    Case(number=1, places=0, expected=NumExcelError),
    Case(number=1, places="0", expected=NumExcelError),
    Case(number=1, places="0", expected=NumExcelError),
    Case(number=1, places=5, expected="00001"),
    Case(number=1, places="5", expected="00001"),
    Case(number=1, places="005", expected="00001"),
    Case(number=1, places="000000000005", expected="00001"),
    Case(number=1, places=10, expected="0000000001"),
    Case(number=1, places=11, expected=NumExcelError),
    Case(number=1, places="11", expected=NumExcelError),
    Case(number=1, places="", expected=ValueExcelError),
)
@parametrize_cases(*[Case(func=i) for i in all_funcs_taking_places])
def test_all_funcs_taking_places(func, number, places, expected):
    assert_equivalent(func(number, places), expected)


@parametrize_cases(Case(number=True), Case(number=False))
@parametrize_cases(Case(places=5), Case(places=True), Case(places=False))
@parametrize_cases(*[Case(func=i) for i in all_funcs_taking_places])
def test_booleans_give_value_error_with_places(func, number, places):
    assert_equivalent(func(number, places), ValueExcelError)


@parametrize_cases(Case(number=True), Case(number=False))
@parametrize_cases(
    Case(places=None, expected=NumExcelError),
    Case(places=0, expected=NumExcelError),
    Case(places=11, expected=NumExcelError),
    Case(places="nonsense", expected=ValueExcelError),
)
@parametrize_cases(*[Case(func=i) for i in all_funcs_taking_places])
def test_booleans_give_other_errors_with_places(func, number, places, expected):
    assert_equivalent(func(number, places), expected)


###################################
#### Test individual functions ####
###################################


@parametrize_cases(
    Case(number=512, expected=NumExcelError),
    Case(number=-513, expected=NumExcelError),
    Case(number=-1, expected="1111111111"),
    Case(number=-2, expected="1111111110"),
    Case(number=-4.2, expected="1111111100"),
    Case(number=-4.8, expected="1111111100"),
)
def test_dec2bin(number, expected):
    assert_equivalent(engineering.DEC2BIN(number), expected)


@parametrize_cases(
    Case(number=2, places=1, expected=NumExcelError),
    Case(number=2, places=5, expected="00010"),
    Case(number=-1, places=1, expected="1111111111"),
)
def test_dec2bin_with_places(number, places, expected):
    assert_equivalent(engineering.DEC2BIN(number, places), expected)


@parametrize_cases(
    Case(number=-1, expected="7777777777"),
    Case(number=-2, expected="7777777776"),
    Case(number=8, expected="10"),
    Case(number=2 ** 29, expected=NumExcelError),
    Case(number=2 ** 29 - 1, expected="3777777777"),
    Case(number=-(2 ** 29) - 1, expected=NumExcelError),
    Case(number=-(2 ** 29), expected="4000000000"),
    Case(number=-4.2, expected="7777777774"),
    Case(number=-4.8, expected="7777777774"),
)
def test_dec2oct(number, expected):
    assert_equivalent(engineering.DEC2OCT(number), expected)


@parametrize_cases(
    Case(number=8, places=1, expected=NumExcelError),
    Case(number=8, places=5, expected="00010"),
    Case(number=-1, places=1, expected="7777777777"),
)
def test_dec2oct_with_places(number, places, expected):
    assert_equivalent(engineering.DEC2OCT(number, places), expected)


@parametrize_cases(
    Case(number=-1, expected="FFFFFFFFFF"),
    Case(number=2 ** 39, expected=NumExcelError),
    Case(number=2 ** 39 - 1, expected="7FFFFFFFFF"),
    Case(number=-(2 ** 39) - 1, expected=NumExcelError),
    Case(number=-(2 ** 39), expected="8000000000"),
    Case(number=-4.2, expected="FFFFFFFFFC"),
    Case(number=-4.8, expected="FFFFFFFFFC"),
)
def test_dec2hex(number, expected):
    assert_equivalent(engineering.DEC2HEX(number), expected)


@parametrize_cases(
    Case(number=16, places=1, expected=NumExcelError),
    Case(number=16, places=5, expected="00010"),
    Case(number=-1, places=1, expected="FFFFFFFFFF"),
)
def test_dec2hex_with_places(number, places, expected):
    assert_equivalent(engineering.DEC2HEX(number, places), expected)


@parametrize_cases(
    Case(number=2, expected=NumExcelError),
    Case(number=10000000, expected="200"),
    Case(number=1000000000, expected="7777777000"),
    Case(number=1111111111, expected="7777777777"),
    Case(number=10000000000, expected=NumExcelError),
)
def test_bin2oct(number, expected):
    assert_equivalent(engineering.BIN2OCT(number), expected)


@parametrize_cases(
    Case(number=1000, places=1, expected=NumExcelError),
    Case(number=1000, places=5, expected="00010"),
    Case(number=1000000000, places=1, expected="7777777000"),
)
def test_bin2oct_with_places(number, places, expected):
    assert_equivalent(engineering.BIN2OCT(number, places), expected)


@parametrize_cases(
    # yes, ___2DEC functions are meant to return numbers and not strings.
    Case(number=2, expected=NumExcelError),
    Case(number=10, expected=2),
    Case(number=1000000000, expected=-512),
    Case(number=1111111111, expected=-1),
    Case(number=10000000000, expected=NumExcelError),
)
def test_bin2dec(number, expected):
    assert_equivalent(engineering.BIN2DEC(number), expected)


@parametrize_cases(
    Case(number=2, expected=NumExcelError),
    Case(number=10, expected="2"),
    Case(number=10000000, expected="80"),
    Case(number=1000000000, expected="FFFFFFFE00"),
    Case(number=10000000000, expected=NumExcelError),
    Case(number=1111111111, expected="FFFFFFFFFF"),
    Case(number="11000000000", expected=NumExcelError),
)
def test_bin2hex(number, expected):
    assert_equivalent(engineering.BIN2HEX(number), expected)


@parametrize_cases(
    Case(number=10000, places=5, expected="00010"),
    Case(number=10000000, places=1, expected=NumExcelError),
    Case(number=100000000, places=5, expected="00100"),
    Case(number=100000000, places=1, expected=NumExcelError),
    Case(number=1000000000, places=1, expected="FFFFFFFE00"),
)
def test_bin2hex_with_places(number, places, expected):
    assert_equivalent(engineering.BIN2HEX(number, places), expected)


@parametrize_cases(
    Case(number=2, expected="10"),
    Case(number=68, expected=NumExcelError),
    Case(number=1000, expected=NumExcelError),
    Case(number=777, expected="111111111"),
    Case(number=7777777000, expected="1000000000"),
)
def test_oc2bin(number, expected):
    assert_equivalent(engineering.OCT2BIN(number), expected)


@parametrize_cases(
    Case(number=2, places=1, expected=NumExcelError),
    Case(number=2, places=5, expected="00010"),
    Case(number=7777777000, places=1, expected="1000000000"),
)
def test_oct2bin_with_places(number, places, expected):
    assert_equivalent(engineering.OCT2BIN(number, places), expected)


@parametrize_cases(
    Case(number=8, expected=NumExcelError),
    Case(number=10, expected=8),
    Case(number=3777777777, expected=536870911),
    Case(number=4000000000, expected=-536870912),
    Case(number=10000000000, expected=NumExcelError),
    Case(number=7777777777, expected=-1),
)
def test_oct2dec(number, expected):
    assert_equivalent(engineering.OCT2DEC(number), expected)


@parametrize_cases(
    Case(number=8, expected=NumExcelError),
    Case(number=10, expected="8"),
    Case(number=10000000000, expected=NumExcelError),
    Case(number=7777777777, expected="FFFFFFFFFF"),
    Case(number=12, expected="A"),
    Case(number=3777777777, expected="1FFFFFFF"),
    Case(number=4000000000, expected="FFE0000000"),
    Case(number=14000000000, expected=NumExcelError),
)
def test_oct2hex(number, expected):
    assert_equivalent(engineering.OCT2HEX(number), expected)


@parametrize_cases(
    Case(number=20, places=1, expected=NumExcelError),
    Case(number=20, places=5, expected="00010"),
    Case(number=4000000000, places=1, expected="FFE0000000"),
)
def test_oct2hex_with_places(number, places, expected):
    assert_equivalent(engineering.OCT2HEX(number, places), expected)


@parametrize_cases(
    Case(number=2, expected="10"),
    Case(number=10, expected="10000"),
    Case(number=200, expected=NumExcelError),
    Case(number="1FF", expected="111111111"),
    Case(number="1e3", expected="111100011"),
    Case(number="1e0", expected="111100000"),
    Case(number="FFFFFFFDFF", expected=NumExcelError),
    Case(number="FFFFFFFE00", expected="1000000000"),
)
def test_hex2bin(number, expected):
    assert_equivalent(engineering.HEX2BIN(number), expected)


@parametrize_cases(
    Case(number=2, places=1, expected=NumExcelError),
    Case(number=2, places=5, expected="00010"),
    Case(number="FFFFFFFE00", places=1, expected="1000000000"),
)
def test_hex2bin_with_places(number, places, expected):
    assert_equivalent(engineering.HEX2BIN(number, places), expected)


@parametrize_cases(
    Case(number=8, expected="10"),
    Case(number=10, expected="20"),
    Case(number="A", expected="12"),
    Case(number="1e3", expected="743"),
    Case(number=20000000, expected=NumExcelError),
    Case(number="1FFFFFFF", expected="3777777777"),
    Case(number="FFDFFFFFFF", expected=NumExcelError),
    Case(number="Ffe0000000", expected="4000000000"),
)
def test_hex2oct(number, expected):
    assert_equivalent(engineering.HEX2OCT(number), expected)


@parametrize_cases(
    Case(number=8, places=1, expected=NumExcelError),
    Case(number=8, places=5, expected="00010"),
    Case(number="FFFFFFFE00", places=1, expected="7777777000"),
)
def test_hex2oct_with_places(number, places, expected):
    assert_equivalent(engineering.HEX2OCT(number, places), expected)


@parametrize_cases(
    Case(number=10, expected=16),
    Case(number=10000000000, expected=NumExcelError),
    Case(number="FFFFffffff", expected=-1),
    Case(number=8000000000, expected=-549755813888),
)
def test_hex2dec(number, expected):
    assert_equivalent(engineering.HEX2DEC(number), expected)
