import datetime
import yearfrac
from dateutil.relativedelta import relativedelta

from . import utils, xl, xlerrors, func_xltypes

# Testing hook.
now = datetime.datetime.now


@xl.register()
@xl.validate_args
def DATE(
        year: func_xltypes.XlNumber,
        month: func_xltypes.XlNumber,
        day: func_xltypes.XlNumber
) -> func_xltypes.XlDateTime:
    """Returns the sequential serial number that represents a particular date.

    https://support.office.com/en-us/article/
        date-function-e36c0c8c-4104-49da-ab83-82328b832349
    """
    if not (0 < year < 9999):
        raise xlerrors.NumExcelError(
            f'Year must be between 1 and 9999, got {year}')

    if year < 1900:
        year = 1900 + year

    # Excel starts counting at 1 and today is inclusive, thus +2
    delta = relativedelta(
        years=year - 1900, months=int(month) - 1, days=int(day) - 1)
    result = utils.EXCEL_EPOCH + delta

    if result <= utils.EXCEL_EPOCH:
        raise xlerrors.NumExcelError(
            f"Date result before {utils.EXCEL_EPOCH}")

    return result


@xl.register()
def TODAY() -> func_xltypes.XlDateTime:
    """Returns the serial number of the current date.

    https://support.office.com/en-us/article/
        today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
    """
    return now()


@xl.register()
@xl.validate_args
def YEARFRAC(
        start_date: func_xltypes.XlDateTime,
        end_date: func_xltypes.XlDateTime,
        basis: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Returns the fraction of the year represented by the number of whole
    days between two dates.

    https://support.office.com/en-us/article/
        yearfrac-function-3844141e-c76d-4143-82b6-208454ddc6a8
    """
    if start_date < utils.EXCEL_EPOCH:
        raise xlerrors.ValueExcelError(
            f'start_date {start_date} must be after {utils.EXCEL_EPOCH}')

    if end_date < utils.EXCEL_EPOCH:
        raise xlerrors.ValueExcelError(
            f'start_date {start_date} must be after {utils.EXCEL_EPOCH}')

    # Switch dates if start_date > end_date
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    # Get Python internal types.
    start_date, end_date = start_date.value, end_date.value

    if basis == 0:  # US 30/360
        return yearfrac.yearfrac(start_date, end_date, '30e360_matu')
    elif basis == 1:  # Actual/actual
        return yearfrac.yearfrac(start_date, end_date, 'act_afb')
    elif basis == 2:  # Actual/360
        return (end_date - start_date).days / 360
    elif basis == 3:  # Actual/365
        return (end_date - start_date).days / 365
    elif basis == 4:  # Eurobond 30/360
        return yearfrac.yearfrac(start_date, end_date, '30e360')

    raise xlerrors.ValueExcelError(
        f'basis must be 0, 1, 2, 3 or 4, got {basis}')
