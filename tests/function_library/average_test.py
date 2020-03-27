
# Excel reference: https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Average
from koala_xlcalculator.koala_types import XLRange
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestAverage(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/AVERAGE.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_average(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        average_result_00 = Average.average(range_00)
        result_00 = 2.5
        self.assertEqual(result_00, average_result_00)

        range_01 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        average_result_01 = Average.average(range_01)
        result_01 = 2.5
        self.assertEqual(result_01, average_result_01)

        range_02 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        average_result_02 = Average.average(range_01, 5)
        result_02 = 3
        self.assertEqual(result_02, average_result_02)


    def test_average_evaluation(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )
