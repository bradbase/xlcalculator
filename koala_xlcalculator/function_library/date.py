
# Excel reference: https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349


from datetime import datetime, date

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from ..koala_types import XLCell


class xDate(KoalaBaseFunction):
    """"""

    @staticmethod
    def xdate(year, month, day):
        """"""

        if isinstance(year, XLCell):
            year = year.value

        if isinstance(month, XLCell):
            month = month.value

        if isinstance(day, XLCell):
            day = day.value

        if type(year) != int:
            return ExcelError("#VALUE!", '%s is not an integer' % str(year))

        if type(month) != int:
            return ExcelError("#VALUE!", '%s is not an integer' % str(month))

        if type(day) != int:
            return ExcelError("#VALUE!", '%s is not an integer' % str(day))

        if year < 0 or year > 9999:
            return ExcelError("#VALUE!", 'Year must be between 1 and 9999, instead %s' % str(year))

        if year < 1900:
            year = 1900 + year

        year, month, day = xDate.normalize_year(year, month, day) # taking into account negative month and day values

        date_0 = datetime(1900, 1, 1)
        date = datetime(year, month, day)

        result = (datetime(year, month, day) - date_0).days + 2

        if result <= 0:
            return ExcelError("#VALUE!", "Date result is negative")

        else:
            return result

    @staticmethod
    def is_number(excel_date):
        """"""
        return isinstance(excel_date, (int, float))


    @staticmethod
    def normalize_year(y, m, d):
        if m <= 0:
            y -= int(abs(m) / 12 + 1)
            m = 12 - (abs(m) % 12)
            xDate.normalize_year(y, m, d)
        elif m > 12:
            y += int(m / 12)
            m = m % 12

        if d <= 0:
            d += xDate.get_max_days_in_month(m, y)
            m -= 1
            y, m, d = xDate.normalize_year(y, m, d)

        else:
            if m in (4, 6, 9, 11) and d > 30:
                m += 1
                d -= 30
                y, m, d = xDate.normalize_year(y, m, d)
            elif m == 2:
                if (xDate.is_leap_year(y)) and d > 29:
                    m += 1
                    d -= 29
                    y, m, d = xDate.normalize_year(y, m, d)
                elif (not xDate.is_leap_year(y)) and d > 28:
                    m += 1
                    d -= 28
                    y, m, d = xDate.normalize_year(y, m, d)
            elif d > 31:
                m += 1
                d -= 31
                y, m, d = xDate.normalize_year(y, m, d)

        return (y, m, d)


    @staticmethod
    def get_max_days_in_month(month, year):
        if not xDate.is_number(year) or not xDate.is_number(month):
            raise TypeError("All inputs must be a number")
        if year <= 0 or month <= 0:
            raise TypeError("All inputs must be strictly positive")

        if month in (4, 6, 9, 11):
            return 30
        elif month == 2:
            if xDate.is_leap_year(year):
                return 29
            else:
                return 28
        else:
            return 31


    @staticmethod
    def is_leap_year(year):
        if not xDate.is_number(year):
            raise TypeError("%s must be a number" % str(year))
        if year <= 0:
            raise TypeError("%s must be strictly positive" % str(year))

        # Watch out, 1900 is a leap according to Excel => https://support.microsoft.com/en-us/kb/214326
        return (year % 4 == 0 and year % 100 != 0 or year % 400 == 0) or year == 1900


    @staticmethod
    def date_from_int(nb):
        if not xDate.is_number(nb):
            raise TypeError("%s is not a number" % str(nb))

        # origin of the Excel date system
        current_year = 1900
        current_month = 0
        current_day = 0

        while(nb > 0):
            if not xDate.is_leap_year(current_year) and nb > 365:
                current_year += 1
                nb -= 365
            elif xDate.is_leap_year(current_year) and nb > 366:
                current_year += 1
                nb -= 366
            else:
                current_month += 1
                max_days = xDate.get_max_days_in_month(current_month, current_year)

                if nb > max_days:
                    nb -= max_days
                else:
                    current_day = nb
                    nb = 0

        return (current_year, current_month, current_day)
