
# Excel reference: https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc

import unittest

from xlfunctions import Choose

from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

class TestChoose(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/CHOOSE.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value_00 = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value_00 )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value_01 = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( excel_value, value_01 )


    def test_evaluation_A3(self):
        excel_value = [[1, 2, 3]]
        value_00 = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual( excel_value, value_00.values.tolist() )
