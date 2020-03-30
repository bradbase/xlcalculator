
# Excel reference: https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8


from datetime import date

from .excel_lib import KoalaBaseFunction
from ..exceptions import ExcelError
from .date import xDate
from ..koala_types import XLCell

class Yearfrac(KoalaBaseFunction):
    """"""

    @staticmethod
    def yearfrac(start_date, end_date, basis = 0):
        """"""

        def actual_nb_days_ISDA(start, end): # needed to separate days_in_leap_year from days_not_leap_year
            """"""

            y1, m1, d1 = start
            y2, m2, d2 = end

            days_in_leap_year = 0
            days_not_in_leap_year = 0

            year_range = list(range(y1, y2 + 1))

            for y in year_range:

                if y == y1 and y == y2:
                    nb_days = date(y2, m2, d2) - date(y1, m1, d1)

                elif y == y1:
                    nb_days = date(y1 + 1, 1, 1) - date(y1, m1, d1)

                elif y == y2:
                    nb_days = date(y2, m2, d2) - date(y2, 1, 1)

                else:
                    nb_days = 366 if xDate.is_leap_year(y) else 365

                if xDate.is_leap_year(y):
                    days_in_leap_year += nb_days

                else:
                    days_not_in_leap_year += nb_days

            return (days_not_in_leap_year, days_in_leap_year)


        def actual_nb_days_AFB_alter(start, end): # http://svn.finmath.net/finmath%20lib/trunk/src/main/java/net/finmath/time/daycount/DayCountConvention_ACT_ACT_YEARFRAC.java
            """"""

            y1, m1, d1 = start
            y2, m2, d2 = end

            delta = date(*end) - date(*start)

            if delta.days <= 365:
                if xDate.is_leap_year(y1) and xDate.is_leap_year(y2):
                    denom = 366

                elif xDate.is_leap_year(y1) and date(y1, m1, d1) <= date(y1, 2, 29):
                    denom = 366

                elif xDate.is_leap_year(y2) and date(y2, m2, d2) >= date(y2, 2, 29):
                    denom = 366

                else:
                    denom = 365
            else:
                year_range = list(range(y1, y2 + 1))
                nb = 0

                for y in year_range:
                    nb += 366 if xDate.is_leap_year(y) else 365

                denom = nb / len(year_range)

            return delta / denom

        if isinstance(start_date, XLCell):
            start_date = start_date.value

        if isinstance(end_date, XLCell):
            end_date = end_date.value

        if not xDate.is_number(start_date):
            return ExcelError('#VALUE!', 'start_date {} must be a number. You supplied {}'.format(str(start_date), type(start_date)))

        if not xDate.is_number(end_date):
            return ExcelError('#VALUE!', 'end_date %s must be number' % str(end_date))

        if start_date < 0:
            return ExcelError('#VALUE!', 'start_date %s must be positive' % str(start_date))

        if end_date < 0:
            return ExcelError('#VALUE!', 'end_date %s must be positive' % str(end_date))

        if start_date > end_date: # switch dates if start_date > end_date
            temp = end_date
            end_date = start_date
            start_date = temp

        y1, m1, d1 = xDate.date_from_int(start_date)
        y2, m2, d2 = xDate.date_from_int(end_date)

        if basis == 0: # US 30/360
            d2 = 30 if d2 == 31 and (d1 == 31 or d1 == 30) else min(d2, 31)
            d1 = 30 if d1 == 31 else d1

            count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
            result = count / 360

        elif basis == 1: # Actual/actual
            result = actual_nb_days_AFB_alter((y1, m1, d1), (y2, m2, d2)).seconds/60/60/24

        elif basis == 2: # Actual/360
            result = (end_date - start_date) / 360

        elif basis == 3: # Actual/365
            result = (end_date - start_date) / 365

        elif basis == 4: # Eurobond 30/360
            d2 = 30 if d2 == 31 else d2
            d1 = 30 if d1 == 31 else d1

            count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
            result = count / 360

        else:
            return ExcelError('#VALUE!', '%d must be 0, 1, 2, 3 or 4' % basis)

        return result
