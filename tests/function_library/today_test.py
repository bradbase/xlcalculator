
# Excel reference: https://support.office.com/en-us/article/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9

import unittest
from datetime import datetime

from xlfunctions import Today

"""
We can't eval this function in a test as the date will need to move.
"""

class TestToday(unittest.TestCase):

    EXCEL_EPOCH = datetime.strptime("1900-01-01", '%Y-%m-%d').date()
    reference_date = datetime.today().date()
    days_since_epoch = reference_date - EXCEL_EPOCH
    todays_ordinal = days_since_epoch.days + 2


    def test_positive_integers(self):
        self.assertEqual(Today.today(), self.todays_ordinal)
