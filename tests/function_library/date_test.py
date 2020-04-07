
# Excel reference: https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

import unittest

import pandas as pd

from koala_xlcalculator.function_library import xDate
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class Test_Date(unittest.TestCase):

    # this doesn't work as loading a spreadsheet with dates doesn't load
    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/DATE.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

        # tokenizer.py", line 335, in getTokens
        # self.model = compiler.read_and_parse_archive(r"./tests/resources/DATE - Copy.xlsx")

    # def teardown(self):
    #     pass

    #
    def test_year_must_be_positive(self):
        self.assertIsInstance(xDate.xdate(-1, 1, 1), ExcelError )


    def test_year_must_have_less_than_10000(self):
        self.assertIsInstance(xDate.xdate(10000, 1, 1), ExcelError )


    def test_result_must_be_positive(self):
        self.assertIsInstance(xDate.xdate(1900, 1, -1), ExcelError )


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

    # this doesn't work as loading a spreadsheet with dates doesn't load
    def test_counta_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( excel_value, value )


    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual( excel_value, value )


    def test_evaluation_A4(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A4')
        value = self.evaluator.evaluate('Sheet1!A4')
        self.assertEqual( excel_value, value )


    def test_evaluation_A5(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A5')
        value = self.evaluator.evaluate('Sheet1!A5')
        self.assertEqual( excel_value, value )


    def test_evaluation_A6(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A6')
        value = self.evaluator.evaluate('Sheet1!A6')
        self.assertEqual( excel_value, value )


    def test_evaluation_A7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A7')
        value = self.evaluator.evaluate('Sheet1!A7')
        self.assertEqual( excel_value, value )


    def test_evaluation_A9(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A9')
        value = self.evaluator.evaluate('Sheet1!A9')
        self.assertEqual( excel_value, value )


    def test_evaluation_A10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual( excel_value, value )


    def test_evaluation_A11(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A11')
        value = self.evaluator.evaluate('Sheet1!A11')
        self.assertEqual( excel_value, value )


    def test_evaluation_A12(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A12')
        value = self.evaluator.evaluate('Sheet1!A12')
        self.assertEqual( excel_value, value.value )


    def test_evaluation_A13(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A13')
        value = self.evaluator.evaluate('Sheet1!A13')
        self.assertEqual( excel_value, value.value )


    def test_evaluation_A14(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A14')
        value = self.evaluator.evaluate('Sheet1!A14')
        self.assertEqual( excel_value, value.value )


    def test_evaluation_A15(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A15')
        value = self.evaluator.evaluate('Sheet1!A15')
        self.assertEqual( excel_value, value )


    def test_evaluation_A16(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A16')
        value = self.evaluator.evaluate('Sheet1!A16')
        self.assertEqual( excel_value, value )


    def test_evaluation_A17(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A17')
        value = self.evaluator.evaluate('Sheet1!A17')
        self.assertEqual( excel_value, value )


    def test_evaluation_A18(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A18')
        value = self.evaluator.evaluate('Sheet1!A18')
        self.assertEqual( excel_value, value )


    def test_evaluation_A19(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A19')
        value = self.evaluator.evaluate('Sheet1!A19')
        self.assertEqual( excel_value, value )


    def test_evaluation_A25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A25')
        value = self.evaluator.evaluate('Sheet1!A25')
        self.assertEqual( excel_value, value )


    def test_evaluation_B25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B25')
        value = self.evaluator.evaluate('Sheet1!B25')
        self.assertEqual( excel_value, value )


    def test_evaluation_C25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C25')
        value = self.evaluator.evaluate('Sheet1!C25')
        self.assertEqual( excel_value, value )


    def test_evaluation_D25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D25')
        value = self.evaluator.evaluate('Sheet1!D25')
        self.assertEqual( excel_value, value )


    def test_evaluation_E25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!E25')
        value = self.evaluator.evaluate('Sheet1!E25')
        self.assertEqual( excel_value, value )


    def test_evaluation_F25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!F25')
        value = self.evaluator.evaluate('Sheet1!F25')
        self.assertEqual( excel_value, value )
