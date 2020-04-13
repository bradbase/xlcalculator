
# Excel reference: https://support.office.com/en-us/article/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2

import unittest

import pandas as pd

from xlcalculator.function_library import Concat
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

class TestConcat(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/CONCAT.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_concat(self):

        concat_result_00 = Concat.concat("SPAM", " ", "SPAM", " ", "SPAM", " ", "SPAM")
        result_00 = "SPAM SPAM SPAM SPAM"
        self.assertTrue(result_00, concat_result_00)

        concat_result_01 = Concat.concat("SPAM", " ", pd.DataFrame([[1, 2],[3, 4]]), " ", "SPAM", " ", "SPAM")
        result_01 = "SPAM 1234 SPAM SPAM"
        self.assertTrue(result_01, concat_result_01)

        concat_result_02 = Concat.concat("SPAM", "SPAM", "SPAM")
        result_02 = "SPAMSPAMSPAM"
        self.assertTrue(result_02, concat_result_02)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( excel_value, value )


    def test_evaluation_A2(self):
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
