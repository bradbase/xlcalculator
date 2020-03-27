
# Excel reference: https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Choose
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator

class TestChoose(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/CHOOSE.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_choose(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        range_01 = pd.DataFrame([[2, 1],[3, 4]])
        range_02 = pd.DataFrame([[1, 2],[4, 3]])
        choose_result_00 = Choose.choose('2', range_00, range_01, range_02)
        result_00 = pd.DataFrame([[2, 1],[3, 4]])
        self.assertTrue(result_00.equals(choose_result_00))


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value_00 = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value_00 )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value_01 = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( excel_value, value_01 )


    @unittest.skip("Range isn't being tokenised properly in choose.")
    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value_00 = self.evaluator.evaluate('Sheet1!A3')
        self.assertTrue( excel_value.equals(value_00) )
