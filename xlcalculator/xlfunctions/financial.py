
from typing import Tuple

import pandas as pd
import numpy_financial as npf
from scipy.optimize import newton

from . import xl, xlerrors, func_xltypes


@xl.register()
@xl.validate_args
def IRR(
        values: func_xltypes.XlArray,
        guess: func_xltypes.XlNumber = None
) -> func_xltypes.XlNumber:
    """Returns the internal rate of return for a series of cash flows

    https://support.office.com/en-us/article/
        irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc
    """
    # `guess` is not used, but unnecessary, since it is a pure perforamnce
    # optimization.
    return npf.irr(xl.flatten(values))


@xl.register()
@xl.validate_args
def NPV(
        rate: func_xltypes.XlNumber,
        *values: Tuple[func_xltypes.XlNumber],
) -> func_xltypes.XlNumber:
    """Calculates the net present value of an investment by using a discount
    rate and a series of future payments (negative values) and income
    (positive values).

    https://support.office.com/en-us/article/
        npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568
    """
    if not len(values):
        raise xlerrors.ValueExcelError('value1 is required')

    cashflow = [float(value) for value in values]
    rate = float(rate)

    if xl.COMPATIBILITY == 'PYTHON':
        return npf.npv(rate, cashflow)

    return sum([
        val * (1 + rate) ** - (i + 1)
        for (i, val) in enumerate(cashflow)
    ])


@xl.register()
def PMT(
        rate: func_xltypes.XlNumber,
        nper: func_xltypes.XlNumber,
        pv: func_xltypes.XlNumber,
        fv: func_xltypes.XlNumber = 0,
        type: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Calculates the payment for a loan based on constant payments and
    a constant interest rate.

    https://support.office.com/en-us/article/
        pmt-function-0214da64-9a63-4996-bc20-214433fa6441
    """
    # WARNING fv & type not used yet - both are assumed to be their
    #         defaults (0)
    # fv = args[3]
    # type = args[4]

    if xl.COMPATIBILITY == 'PYTHON':
        when = 'end'
        if type != 0:
            when = 'begin'
        return float(npf.pmt(
            float(rate), float(nper), float(pv), fv=float(fv), when=when))

    # return -pv * rate / (1 - power(1 + rate, -nper))
    return float(npf.pmt(
        float(rate), float(nper), float(pv), fv=float(fv), when='end'))


@xl.register()
@xl.validate_args
def PV(
        rate: func_xltypes.XlNumber,
        nper: func_xltypes.XlNumber,
        pmt: func_xltypes.XlNumber,
        fv: func_xltypes.XlNumber = 0,
        type: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """PV, one of the financial functions, calculates the present value of a
    loan or an investment, based on a constant interest rate.

    https://support.office.com/en-us/article/
        pv-function-23879d31-0e02-4321-be01-da16e8168cbd
    """

    return npf.pv(float(rate),
                  float(nper),
                  float(pmt),
                  fv=float(fv),
                  when=int(type))


@xl.register()
def SLN(
        cost: func_xltypes.XlNumber,
        salvage: func_xltypes.XlNumber,
        life: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the straight-line depreciation of an asset for one period.

    https://support.office.com/en-us/article/
        sln-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8
    """
    return (cost - salvage) / life


@xl.register()
@xl.validate_args
def VDB(
        cost: func_xltypes.XlNumber,
        salvage: func_xltypes.XlNumber,
        life: func_xltypes.XlNumber,
        start_period: func_xltypes.XlNumber,
        end_period: func_xltypes.XlNumber,
        factor: func_xltypes.XlNumber = 2,
        no_switch: func_xltypes.XlBoolean = False
) -> func_xltypes.XlNumber:
    """Returns the depreciation of an asset for any period you specify.

    https://support.office.com/en-us/article/
        vdb-function-dde4e207-f3fa-488d-91d2-66d55e861d73
    """
    start_period = start_period
    end_period = end_period
    sln_depr = SLN(cost, salvage, life)
    depr_rate = factor / life
    acc_depr = 0
    depr = 0
    switch_to_sln = False
    sln_depr = 0
    result = 0
    start_life = 0
    # Handle cases when life is not an integer.
    if life % 1 > 0:
        end_life = int(life + 1)

    else:
        end_life = int(life)
    periods = list(range(start_life, end_life))

    if int(start_period) != start_period:
        delta_start = abs(int(start_period) - start_period)
        depr = (cost - acc_depr) * depr_rate * delta_start
        acc_depr += depr
        start_life = 1
        periods = [x + 0.5 for x in periods]

    for index, current_year in enumerate(periods):

        if not no_switch:  # no_switch = False (Default Case)
            if switch_to_sln:
                depr = sln_depr

            else:
                depr = (cost - acc_depr) * depr_rate
                acc_depr += depr

                temp_sln_depr = SLN(cost, salvage, life)

                if depr < temp_sln_depr:
                    switch_to_sln = True
                    fixed_remaining_years = life - current_year - 1
                    fixed_remaining_cost = cost - acc_depr

                    # We need to check future sln: current depr should never
                    # be smaller than sln to come.
                    sln_depr = SLN(
                        fixed_remaining_cost, salvage, fixed_remaining_years)

                    # if it's the case, we switch to sln earlier than the
                    # regular case
                    if sln_depr > depr:
                        # cancel what has been done
                        acc_depr -= depr
                        fixed_remaining_years += 1
                        fixed_remaining_cost = cost - acc_depr

                        # recalculate depreciation
                        sln_depr = SLN(
                            fixed_remaining_cost, salvage,
                            fixed_remaining_years)
                        depr = sln_depr
                        acc_depr += depr

        else:  # no_switch = True
            depr = (cost - acc_depr) * depr_rate
            acc_depr += depr

        delta_start = abs(current_year - start_period)

        if delta_start < 1 and delta_start != 0:
            result += depr * (1 - delta_start)

        elif current_year >= start_period and current_year < end_period:

            delta_end = abs(end_period - current_year)

            if delta_end < 1 and delta_end != 0:
                result += depr * delta_end

            else:
                result += depr

    return result


def _xnpv(rate, values, dates):
    if rate <= -1.0:
        return float('inf')

    return sum([value / ((1.0 + rate)**((date - dates[0]) / 365))
               for value, date in zip(values, dates)])


def _xirr(values, dates, guess=None):
    try:
        return newton(lambda r: _xnpv(r, values, dates), guess, maxiter=100)

    except (RuntimeError, FloatingPointError):
        raise xlerrors.NumExcelError('XIRR did not converge')


@xl.register()
@xl.validate_args
def XIRR(
        values: func_xltypes.XlArray,
        dates: func_xltypes.XlArray,
        guess: func_xltypes.XlNumber = 0.1
) -> func_xltypes.XlNumber:
    """Returns the internal rate of return for a schedule of cash flows that
        is not necessarily periodic.

    https://support.microsoft.com/en-us/office/
        xirr-function-de1242ec-6477-445b-b11b-a303ad9adc9d

    Algorithm found on stackoverflow:
    https://stackoverflow.com/questions/63797804/
        python-irr-function-giving-different-result-than-excel-xirr

    From MS, Newton's method is used to optimize:
    https://docs.microsoft.com/en-us/office/troubleshoot/excel/
        algorithm-of-xirr-funcation
    """

    values = values.flatten(func_xltypes.Number, None)
    dates = dates.flatten(func_xltypes.DateTime, None)
    # need to cast dates and guess to Python types else optimizer complains
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

    # Sort dataframe by date
    series = series.sort_values('dates', ascending=True)
    series['values'] = series['values'].astype('float')

    # Create separate lists for values and dates
    series_values = list(series['values'])
    series_dates = list(series['dates'])

    # Calculate IRR
    return _xirr(series_values, series_dates, guess)


@xl.register()
@xl.validate_args
def XNPV(
        rate: func_xltypes.XlNumber,
        values: func_xltypes.XlArray,
        dates: func_xltypes.XlArray,
) -> func_xltypes.XlNumber:
    """Returns the net present value for a schedule of cash flows that
    is not necessarily periodic.

    https://support.microsoft.com/en-us/office/
        xnpv-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7
    """
    values = values.flatten(func_xltypes.Number, None)
    dates = dates.flatten(func_xltypes.DateTime, None)

    # TODO: Ignore non numeric cells and boolean cells.
    if len(values) != len(dates):
        raise xlerrors.NumExcelError(
            f'`values` range must be the same length as `dates` range '
            f'in XNPV, {len(values)} != {len(dates)}')

    return _xnpv(rate, values, dates)
