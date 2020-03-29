
# Excel reference: https://support.office.com/en-us/article/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Sumproduct
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator.koala_types import XLRange
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestSumProduct(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/SUMPRODUCT.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_ranges_with_different_sizes(self):
        range1 = XLRange("Sheet1!A1:A3", "Sheet1!A1:A3", value = pd.DataFrame([[1], [10], [3]]))
        range2 = XLRange("Sheet1!A1:A3", "Sheet1!A1:A3", value = pd.DataFrame([[3], [3], [1], [2]]))

        with self.assertRaises(ExcelError):
            Sumproduct.sumproduct(range1, range2)


    def test_regular(self):
        range1 = XLRange("Sheet1!A1:A3", "Sheet1!A1:A3", value = pd.DataFrame([[1], [10], [3]]))
        range2 = XLRange("Sheet1!A1:A3", "Sheet1!A1:A3", value = pd.DataFrame([[3], [1], [2]]))

        self.assertEqual(Sumproduct.sumproduct(range1, range2), 19)


    def test_evaluation_D7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D7')
        value = self.evaluator.evaluate('Sheet1!D7')
        self.assertEqual( excel_value, value )
