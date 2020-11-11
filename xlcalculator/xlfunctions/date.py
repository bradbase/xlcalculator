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

    # Excel starts counting at 1 and today is inclusive, thus -2
    delta = relativedelta(
        years=year - 1900, months=int(month) - 1, days=int(day) - 1)
    result = utils.EXCEL_EPOCH + delta

    if result <= utils.EXCEL_EPOCH:
        raise xlerrors.NumExcelError(
            f"Date result before {utils.EXCEL_EPOCH}")

    return result


@xl.register()
@xl.validate_args
def DAY(
        serial_number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the day of a date, represented by a serial number. The day is
    given as an integer ranging from 1 to 31.

    https://support.microsoft.com/en-us/office/
        day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101
    """

    date = utils.number_to_datetime(int(serial_number))
    return int(date.strftime("%d"))


@xl.register()
@xl.validate_args
def MONTH(
        serial_number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the month of a date represented by a serial number. The month
    is given as an integer, ranging from 1 (January) to 12 (December).

    https://support.microsoft.com/en-us/office/
        month-function-579a2881-199b-48b2-ab90-ddba0eba86e8
    """

    date = utils.number_to_datetime(int(serial_number))
    return int(date.strftime("%m"))


@xl.register()
def NOW() -> func_xltypes.XlNumber:
    """Returns the serial number of the current date and time.

    https://support.office.com/en-us/article/
        today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
    """
    return utils.datetime_to_number(now())


@xl.register()
def TODAY() -> func_xltypes.XlNumber:
    """Returns the serial number of the current date.

    https://support.office.com/en-us/article/
        today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
    """
    date_and_time = now()
    date = date_and_time.replace(
        hour=0, minute=0, second=0, microsecond=0)
    return utils.datetime_to_number(date)


@xl.register()
@xl.validate_args
def YEAR(
        serial_number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the year corresponding to a date. The year is returned as an
        integer in the range 1900-9999.

    https://support.microsoft.com/en-us/office/
        year-function-c64f017a-1354-490d-981f-578e8ec8d3b9
    """

    date = utils.number_to_datetime(int(serial_number))

    if (int(date.strftime("%Y")) < int(utils.EXCEL_EPOCH.strftime("%Y"))) \
            or (int(date.strftime("%Y")) > 9999):
        raise xlerrors.ValueExcelError(
            f'year {date.strftime("%Y")} must be after \
            {utils.EXCEL_EPOCH.strftime("%Y")} and before 9999')

    return int(date.strftime("%Y"))


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
