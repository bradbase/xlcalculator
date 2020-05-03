
# Excel reference: https://support.office.com/en-us/article/sln-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8

import unittest

from xlfunctions import SLN

from xlcalculator.xlcalculator_types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator


class TestSLN(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/SLN.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )
