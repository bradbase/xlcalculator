
# Excel reference: https://support.office.com/en-us/article/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89

import unittest

import pandas as pd

from koala_xlcalculator.function_library import xSum
from koala_xlcalculator.koala_types import XLRange
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestSum(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/SUM.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_sum(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        xsum_result_00 = xSum.xsum(range_00)
        result_00 = 10
        self.assertEqual(result_00, xsum_result_00)

        range_01 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        xsum_result_01 = xSum.xsum(range_01)
        result_01 = 10
        self.assertEqual(result_01, xsum_result_01)


    def test_evaluation_A10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual( excel_value, value )
