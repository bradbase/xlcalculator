from __future__ import annotations

import os
import unittest
from dataclasses import dataclass
from decimal import Decimal, ROUND_UP, ROUND_DOWN
from typing import Any, Optional, Callable
from _pytest.mark.structures import MarkDecorator
import pytest

from xlcalculator import xlerrors
from xlcalculator.model import ModelCompiler
from xlcalculator.evaluator import Evaluator


RESOURCE_DIR = os.path.join(os.path.dirname(__file__), 'resources')


def get_resource(filename):
    return os.path.join(RESOURCE_DIR, filename)


@dataclass
class f_token:

    tvalue: str
    ttype: str
    tsubtype: str

    @classmethod
    def from_token(cls, token):
        return cls(token.tvalue, token.ttype, token.tsubtype)

    def __repr__(self):
        return "<{} tvalue: {} ttype: {} tsubtype: {}>".format(
            self.__class__.__name__, self.tvalue, self.ttype, self.tsubtype)

    def __str__(self):
        return self.__repr__()


class XlCalculatorTestCase(unittest.TestCase):

    def assertEqualRounded(self, lhs, rhs, rounding_precision=None):

        if rounding_precision is None:
            lhs_split = str(lhs).split('.')
            rhs_split = str(rhs).split('.')

            if len(lhs_split) > 1:
                len_lhs_after_decimal = len(lhs_split[1])
            else:
                len_lhs_after_decimal = None

            if len(rhs_split) > 1:
                len_rhs_after_decimal = len(rhs_split[1])
            else:
                len_rhs_after_decimal = None

            if len_lhs_after_decimal is None or len_rhs_after_decimal is None:
                return self.assertEqual(round(lhs), round(rhs))

            rounding_precision = min(
                len_lhs_after_decimal, len_rhs_after_decimal)

        precision_mask = "{0:." + str(rounding_precision - 1) + "f}1"
        precision = precision_mask.format(0.0)

        if lhs > rhs:
            lhs_value = Decimal(lhs).quantize(Decimal(
                precision), rounding=ROUND_DOWN)
            rhs_value = Decimal(rhs).quantize(
                Decimal(precision), rounding=ROUND_UP)
        else:
            lhs_value = Decimal(rhs).quantize(
                Decimal(precision), rounding=ROUND_UP)
            rhs_value = Decimal(lhs).quantize(
                Decimal(precision), rounding=ROUND_DOWN)

        return self.assertEqual(lhs_value, rhs_value)

    def assertEqualTruncated(self, lhs, rhs, truncating_places=None):
        lhs_before_dec, lhs_after_dec = str(lhs).split('.')
        rhs_before_dec, rhs_after_dec = str(rhs).split('.')

        if truncating_places is None:
            truncating_places = min(
                len(str(lhs).split('.')[1]), len(str(rhs).split('.')[1]))

        if 'E' in lhs_after_dec:
            lhs_value = float('.'.join((lhs_before_dec, lhs_after_dec)))
        else:
            lhs_value = float('.'.join((
                lhs_before_dec, lhs_after_dec[0:truncating_places])))

        if 'E' in lhs_after_dec:
            rhs_value = float('.'.join((rhs_before_dec, rhs_after_dec)))
        else:
            rhs_value = float('.'.join((
                rhs_before_dec, rhs_after_dec[0:truncating_places])))

        return self.assertAlmostEqual(lhs_value, rhs_value, truncating_places)

    def assertASTNodesEqual(self, lhs, rhs):
        lhs = [f_token.from_token(t) for t in lhs]
        rhs = [f_token.from_token(t) for t in rhs]
        return self.assertEqual(lhs, rhs)


class FunctionalTestCase(XlCalculatorTestCase):

    filename = None

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            get_resource(self.filename))
        self.evaluator = Evaluator(self.model)


class Case:
    """Container for a test case, with optional test ID.
    Attributes:
        __label__: Optional test ID string. Will be displayed for each test
            when running `pytest -v`
        kwargs: Parameters used for the test cases.
    Example:
        Case("some test name", foo=10, bar="some value")
        Case(foo=99, bar="some other value") # no label given
    """

    def __init__(self, __label__: Optional[str] = None, **kwargs: Any):
        """Initializes Case class with label (optional) and kwargs."""
        self.label = __label__
        self.kwargs = kwargs

    def __eq__(self, other: object) -> bool:
        """Return self==value."""
        if not isinstance(other, Case):
            return False

        return (self.label == other.label) and (self.kwargs == other.kwargs)

    def __repr__(self) -> str:
        """Return repr(self)."""
        pairs = [f"{key}={value!r}," for key, value in self.kwargs.items()]
        joined = " ".join(pairs)
        label = f"'{self.label}', " if self.label is not None else ""
        return f"Case({label}{joined})"


def parametrize_cases(*cases: Case) -> MarkDecorator:
    """Decorator wrapper for pytest.mark.parametrize.
    Args:
        *cases:
            One or more Case objects. They must all have the same set of named
            keyword arguments.
    Returns:
        A suitable MarkDecorator instance.
    Example:
        from datetime import datetime, timedelta
        @parametrize_cases(
            Case(
                "forward",
                a=datetime(2001, 12, 12),
                b=datetime(2001, 12, 11),
                expected=timedelta(1),
            ),
            Case(
                "backward",
                a=datetime(2001, 12, 11),
                b=datetime(2001, 12, 12),
                expected=timedelta(-1),
            ),
        )
        def test_timedistance(a, b, expected):
            diff = a - b
            assert diff == expected
    """
    first_case = cases[0]
    first_args = first_case.kwargs.keys()

    for case in cases:
        if first_args != case.kwargs.keys():
            msg = f"Inconsistent parametrization: {first_case!r}, {case!r}"
            raise ValueError(msg)

    argnames = ",".join(first_args)

    argvalues = [tuple(case.kwargs[i] for i in first_args) for case in cases]
    if len(first_args) == 1:
        argvalues = [i[0] for i in argvalues]

    ids = [case.label for case in cases]
    if all(i is None for i in ids):
        return pytest.mark.parametrize(argnames=argnames, argvalues=argvalues)

    return pytest.mark.parametrize(argnames=argnames, argvalues=argvalues, ids=ids)


def assert_equivalent(result, expected, normalize: Optional[Callable] = None):
    if isinstance(expected, type) and issubclass(expected, xlerrors.ExcelError):
        assert isinstance(result, expected), f"Expected {expected!r}, got {result!r}"
    elif normalize:
        assert normalize(result, expected)
    else:
        assert result == expected, f"Expected {expected!r}, got {result!r}"


def workbook_test_cases(filename: str) -> MarkDecorator:
    compiler = ModelCompiler()
    resolved_filename = get_resource(filename)
    model = compiler.read_and_parse_archive(resolved_filename)
    evaluator = Evaluator(model)

    cases = [
        Case(
            f"{filename} {address}",
            sheet_value=evaluator.get_cell_value(address),
            calculated_value=evaluator.evaluate(address)
        )
        for address in model.formulae
    ]
    return parametrize_cases(*cases)
