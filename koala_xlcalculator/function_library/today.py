
# https://support.office.com/en-ie/article/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9


from datetime import datetime

from .excel_lib import KoalaBaseFunction

class Today(KoalaBaseFunction):
    """"""

    def today(self):
        """"""

        reference_date = datetime.today().date()
        days_since_epoch = reference_date - KoalaBaseFunction.EXCEL_EPOCH
        # why +2 ?
        # 1 based from 1900-01-01
        # I think it is "inclusive" / to the _end_ of the day.
        # https://support.office.com/en-us/article/date-function-e36c0c8c-4104-49da-ab83-82328b832349
        """Note: Excel stores dates as sequential serial numbers so that they can be used in calculations.
        January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,447 days after January 1, 1900.
         You will need to change the number format (Format Cells) in order to display a proper date."""
        return days_since_epoch.days + 2
