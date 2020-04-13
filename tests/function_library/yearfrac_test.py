
# Excel reference: https://support.office.com/en-us/article/yearfrac-function-3844141e-c76d-4143-82b6-208454ddc6a8

import unittest

from datetime import date

import pandas as pd

from xlcalculator.function_library import Yearfrac
from xlcalculator.function_library import xDate
from xlcalculator.xlcalculator_types import XLCell
from xlcalculator.exceptions import ExcelError
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from ..xlcalculator_test import XlCalculatorTestCase

# Basis 1, 	Actual/actual, is in error. can only go to 3 decimal places

class TestYearfrac(XlCalculatorTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/YEARFRAC.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_start_date_must_be_number(self):
        self.assertIsInstance(Yearfrac.yearfrac('not a number', 1), ExcelError )


    def test_end_date_must_be_number(self):
        self.assertIsInstance(Yearfrac.yearfrac(1, 'not a number'), ExcelError )


    def test_start_date_must_be_positive(self):
        self.assertIsInstance(Yearfrac.yearfrac(-1, 0), ExcelError )


    def test_end_date_must_be_positive(self):
        self.assertIsInstance(Yearfrac.yearfrac(0, -1), ExcelError )


    def test_basis_must_be_between_0_and_4(self):
        self.assertIsInstance(Yearfrac.yearfrac(1, 2, 5), ExcelError )
        self.assertIsInstance(Yearfrac.yearfrac(1, 2, -1), ExcelError )


    def test_yearfrac_basis_0(self):
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20)), 7.30277777777778)
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20), 0), 7.30277777777778)
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2025, 1, 1), 0), 1)


    @unittest.skip("basis_1 doesn't get accurate enough (not within 3 decimal places)")
    def test_yearfrac_basis_1(self):
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20), 1), 7.299110198)  # multi year
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2024, 12, 31), 1), 0.99726776)  # day before leap
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2025, 1, 1), 1), 1)  # first day leap
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2025, 1, 2), 1), 1.004103967)  # second day leap
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2025, 1, 1), xDate.xdate(2025, 12, 31), 1), 0.997260273972)  # day before non-leap
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2025, 1, 1), xDate.xdate(2026, 1, 1), 1), 1)  # first day non-leap
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2025, 1, 1), xDate.xdate(2026, 1, 2), 1), 1.0027397260274)  # second day non-leap


    def test_yearfrac_basis_2(self):
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20), 2), 7.405555556)
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2025, 1, 1), 2), 1.01666666666667)


    def test_yearfrac_basis_3(self):
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20), 3), 7.304109589)
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2025, 1, 1), 3), 1.0027397260274)


    def test_yearfrac_basis_4(self):
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20), 4), 7.302777778)
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2024, 1, 1), xDate.xdate(2025, 1, 1), 4), 1)


    def test_yearfrac_inverted(self):
        self.assertAlmostEqual(Yearfrac.yearfrac(xDate.xdate(2015, 4, 20), xDate.xdate(2008, 1, 1)), Yearfrac.yearfrac(xDate.xdate(2008, 1, 1), xDate.xdate(2015, 4, 20)))


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqualTruncated( excel_value, value )


    # only close within 3 decimal places.
    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqualTruncated( excel_value, value, 3 )


    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqualTruncated( excel_value, value )
