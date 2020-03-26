
import unittest

import pandas as pd

from koala_xlcalculator.function_library import xDate
from koala_xlcalculator.exceptions import ExcelError


class Test_Date(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_year_must_be_integer(self):
        with self.assertRaises(ExcelError):
            xDate.xdate('2016', 1, 1)


    def test_month_must_be_integer(self):
        with self.assertRaises(ExcelError):
            xDate.xdate(2016, '1', 1)


    def test_day_must_be_integer(self):
        with self.assertRaises(ExcelError):
            xDate.xdate(2016, 1, '1')


    def test_year_must_be_positive(self):
        with self.assertRaises(ExcelError):
            xDate.xdate(-1, 1, 1)


    def test_year_must_have_less_than_10000(self):
        with self.assertRaises(ExcelError):
            xDate.xdate(10000, 1, 1)


    def test_result_must_be_positive(self):
        with self.assertRaises(ExcelError):
            xDate.xdate(1900, 1, -1)


    def test_not_stricly_positive_month_substracts(self):
        self.assertEqual(xDate.xdate(2009, -1, 1), xDate.xdate(2008, 11, 1))


    def test_not_stricly_positive_day_substracts(self):
        self.assertEqual(xDate.xdate(2009, 1, -1), xDate.xdate(2008, 12, 30))


    def test_month_superior_to_12_change_year(self):
        self.assertEqual(xDate.xdate(2009, 14, 1), xDate.xdate(2010, 2, 1))


    def test_day_superior_to_365_change_year(self):
        self.assertEqual(xDate.xdate(2009, 1, 400), xDate.xdate(2010, 2, 4))


    def test_year_for_29_feb(self):
        self.assertEqual(xDate.xdate(2008, 2, 29), 39507)


    def test_year_regular(self):
        self.assertEqual(xDate.xdate(2000, 1, 1), 36526)
        self.assertEqual(xDate.xdate(2008, 11, 3), 39755)
        self.assertEqual(xDate.xdate(2024, 1, 1), 45292)
        self.assertEqual(xDate.xdate(2025, 1, 1), 45658)
        self.assertEqual(xDate.xdate(2026, 1, 1), 46023)
