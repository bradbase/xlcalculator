
# Excel reference: https://support.office.com/en-us/article/max-function-e0012414-9ac8-4b34-9a47-73e662c08098

import unittest

import pandas as pd

from koala_xlcalculator.function_library import xMax
from koala_xlcalculator.koala_types import XLRange
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestMax(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/MAX.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass


    def test_max(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        xsum_result_00 = xMax.xmax(range_00)
        result_00 = 4
        self.assertEqual(result_00, xsum_result_00)

        range_01 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        xsum_result_01 = xMax.xmax(range_01)
        result_01 = 4
        self.assertEqual(result_01, xsum_result_01)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )
