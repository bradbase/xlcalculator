
# Excel reference: https://support.office.com/en-us/article/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89

import unittest

from xlfunctions import xSum

from xlcalculator.xlcalculator_types import XLRange
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator


class TestSum(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/SUM.xlsx")
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_evaluation_A10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual( excel_value, value )
