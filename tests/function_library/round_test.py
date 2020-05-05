
# Excel reference: https://support.office.com/en-us/article/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

import unittest

from xlfunctions import xRound
from xlfunctions.exceptions import ExcelError

from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

class Test_Round(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/ROUND.xlsx")
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


    def test_evaluation_A6(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A6')
        value = self.evaluator.evaluate('Sheet1!A6')
        self.assertEqual( excel_value, value )


    def test_evaluation_A7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A7')
        value = self.evaluator.evaluate('Sheet1!A7')
        self.assertEqual( excel_value, value )
