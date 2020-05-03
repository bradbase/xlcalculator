
# Excel reference: https://support.office.com/en-us/article/rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53

import unittest

from xlfunctions import xRound
from xlfunctions.exceptions import ExcelError

from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

class Test_Rounddown(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/ROUNDDOWN.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_A2(self):
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
