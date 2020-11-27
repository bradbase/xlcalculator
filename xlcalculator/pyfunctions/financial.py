
import datetime
import decimal
from typing import List

import pandas as pd
from scipy.optimize import newton, brentq

from ..xlfunctions import xlerrors, func_xltypes, xl
from . import pyxl


def _xnpv(
        rate: float,
        values: List[decimal.Decimal],
        dates: List[float]
) -> float:
    """XNPV"""
    if rate <= -1:
        return float('inf')
    d0 = dates[0]    # or min(dates)
    try:
        return sum([float(vi) / (1.0 + rate)**((di - d0) / 365)
                   for vi, di in zip(values, dates)])
    except OverflowError as e:  # pragma: no cover
        # If the number we are dealing with gets too large, bail out
        raise RuntimeError from e


def _xirr(
        values: List[decimal.Decimal],
        dates: List[datetime.date],
        guess: float = 0.1
) -> float:
    """XIRR."""
    if all([dt == dates[0] for dt in dates]):
        return 0.0
    try:
        return newton(lambda r: _xnpv(r, values, dates),
                      guess,
                      maxiter=100)
    except (RuntimeError, OverflowError):  # Failed to converge.
        return brentq(lambda r: _xnpv(r, values, dates), -1.0, 1e10)


@pyxl.registerpy()
@xl.validate_args
def XIRR(
        values: func_xltypes.XlArray,
        dates: func_xltypes.XlArray,
        guess: func_xltypes.XlNumber = 0.1
) -> func_xltypes.XlNumber:
    """Returns the internal rate of return for a schedule of cash flows that
        is not necessarily periodic.
    """
    values = values.flatten(func_xltypes.Number, None)
    dates = dates.flatten(func_xltypes.DateTime, None)
    # need to cast dates and guess to Python types else optimizer complains
    # values = [float(value) for value in values]
    dates = [float(date) for date in dates]
    guess = float(guess)

    # TODO: Ignore non numeric cells and boolean cells.
    if len(values) != len(dates):
        raise xlerrors.NumExcelError(
            f'`values` range must be the same length as `dates` range '
            f'in XIRR, {len(values)} != {len(dates)}')

    series = pd.DataFrame({"dates": dates, "values": values})

    # Filter all rows with 0 cashflows
    series = series[series['values'] != 0]

    # While mathematically defined, all positive or all negative values do not
    # really represent a valid scenario, so returning None is better than
    # infinity.
    if all(series['values'] >= 0):
        return None  # Infinity
    if all(series['values'] <= 0):
        return None  # -Infinity

    # Sort dataframe by date
    series = series.sort_values('dates', ascending=True)
    series['values'] = series['values'].astype('float')

    # Create separate lists for values and dates
    moneyIn = list(series['values'])
    dates = list(series['dates'])

    try:

        return _xirr(moneyIn, dates, guess)
    except RuntimeError:
        # The XIRR function can raise a runtime error if it cannot converge to
        # a value.
        return None
    except (ValueError, decimal.InvalidOperation):
        # Happened at least once that the return was so high, xirr errd out.
        return None
