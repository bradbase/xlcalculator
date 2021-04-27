from dataclasses import dataclass
from typing import Sequence, Callable, Union

import xlwings
import pytest
import hypothesis

from xlcalculator import xlerrors
from tests.testing import assert_equivalent
from tests.conftest import CONFIG


@pytest.fixture(scope="session")
def excel_app():
    app = xlwings.App(add_book=False, visible=True)
    try:
        yield app
    finally:
        app.kill()


@pytest.fixture
def excel_workbook(excel_app):
    wb = excel_app.books.add()
    try:
        yield wb
    finally:
        wb.close()


error_lookup = {
    # xlwings does not (yet) let us inspect a cell to determine what kind of error it
    # holds, so we have to do a workaround where we use Excel's ERROR.TYPE function, which
    # returns an integer depending on what kind of error the given cell has.
    1: xlerrors.NullExcelError,
    2: xlerrors.DivZeroExcelError,
    3: xlerrors.ValueExcelError,
    4: xlerrors.RefExcelError,
    5: xlerrors.NameExcelError,
    6: xlerrors.NumExcelError,
    7: xlerrors.NaExcelError,
}


@dataclass
class FormulaTestingEnvironment:
    """Given an xlwings workbook, a function, and the names of the arguments of that
    function, initializes cells on the workbook so that the function can be tested against
    its Excel formula equivalent.

    For example,

    >>> env = FormulaTestingEnvironment(
    ...     wb=<some blank workbook>, formula=DEC2BIN, argnames=["number", "places"]
    ... )

    We can then set the "number" argument via the `set_args` method:

    >>> env.set_args(number=123, places=9)

    And get the formula result back from the live Excel workbook:

    >>> env.value
    '001111011'

    At the moment, it supports only single-cell, named arguments.
    """

    wb: xlwings.main.Book
    formula: Callable
    argnames: Sequence[str]

    def __post_init__(self):
        sheet = self.wb.sheets["Sheet1"]

        arg_addresses = {name: f"A{idx + 1}" for idx, name in enumerate(self.argnames)}
        self.args = {name: sheet[address] for name, address in arg_addresses.items()}

        args_string = ",".join(arg_addresses.values()).rstrip(",")
        formula_string = f"={self.formula_name}({args_string})"
        formula_address = "B1"
        self.formula_cell = sheet[formula_address]
        self.formula_cell.value = formula_string

        is_error_address = "C1"
        self.is_error_cell = sheet[is_error_address]
        self.is_error_cell.value = f"=ISERROR({formula_address})"

        error_type_address = "D1"
        self.error_type_cell = sheet[error_type_address]
        self.error_type_cell.value = f"=ERROR.TYPE({formula_address})"

    @property
    def formula_name(self):
        return self.formula.__name__

    @property
    def is_error(self):
        return self.is_error_cell.value

    @property
    def error_type(self):
        return error_lookup[self.error_type_cell.value]

    @property
    # awkward workaround because xlwings doesn't report the error code directly
    def value(self):
        if self.is_error:
            return self.error_type
        else:
            return self.formula_cell.value

    def set_args(self, **kwargs):
        for name, value in kwargs.items():
            if isinstance(value, str):
                value = (
                    "'" + value
                )  # prevent 1e3 being interpreted as scientific notation
            self.args[name].value = value


# The object returned by hypothesis.given does not let us inspect the kwargs that were
# used in its construction. We do need to access them in fuzz_scalars, so we have to make
# this wrapped version that saves kwargs as attributes on the return value.
_given = hypothesis.given


def given(*args, **kwargs):
    res = _given(*args, **kwargs)
    setattr(res, "given_args", args)
    setattr(res, "given_kwargs", kwargs)
    return res


def fuzz_scalars(wb, formula, values, settings=None):
    """Compare Python implementation of formula to actual Excel formula.

    Tests a Python implementation of an Excel formula on arbitrary hypothesis-generated
    values, by comparing it to the results of that formula in a running Excel instance.
    i.e. use Excel as a live "oracle". Where the results disagree, the assert fails.

    Parameters:

    wb: an xlwings Workbook
    formula: a function that is a Python implementation of an Excel formula
    values: the result of a hypothesis.given invocation (would normally be a decorator)
    settings: optional hypothesis.settings object (e.g. to override max_examples)
    """
    
    if not hasattr(values, "given_kwargs"):
        raise ValueError(
            f"You can't use the `given` function from hypothesis; you need to use the patched version defined in {given.__module__}.py"
        )

    argnames = list(values.given_kwargs.keys())
    env = FormulaTestingEnvironment(wb, formula, argnames)
    
    # workaround limitation / design choice in hypothesis regarding reuse of
    # function-scoped fixtures; see this github issue:
    # https://github.com/HypothesisWorks/hypothesis/issues/2731
    def inner(**kwargs_from_hypothesis):
        python_result = env.formula(**kwargs_from_hypothesis)
        env.set_args(**kwargs_from_hypothesis)
        excel_result = env.value

        assert_equivalent(result=python_result, expected=excel_result)

    settings = settings or hypothesis.settings(
        deadline=None, max_examples=CONFIG["max-examples"]
    )
    func = settings(values(inner))
    func()
