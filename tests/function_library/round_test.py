
# Excel reference: https://support.office.com/en-us/article/round-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

import unittest

import pandas as pd

from xlcalculator.function_library import xRound
from xlcalculator.exceptions import ExcelError
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

class Test_Round(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/ROUND.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_nb_must_be_number(self):
        self.assertIsInstance(xRound.xround('er', 1), ExcelError )


    def test_nb_digits_must_be_number(self):
        self.assertIsInstance(xRound.xround(2.323, 'ze'), ExcelError )


    def test_positive_number_of_digits(self):
        self.assertEqual(xRound.xround(2.675, 2), 2.68)


    def test_negative_number_of_digits(self):
        self.assertEqual(xRound.xround(2352.67, -2), 2400)


    def test_round(self):
        self.assertEqual(xRound.xround(2.15, 1), 2.2)
        self.assertEqual(xRound.xround(2.149, 1), 2.1)
        self.assertEqual(xRound.xround(-1.475, 2), -1.48)
        self.assertEqual(xRound.xround(21.5, -1), 20)
        self.assertEqual(xRound.xround(626.3,-3), 1000)
        self.assertEqual(xRound.xround(1.98,-1), 0)
        self.assertEqual(xRound.xround(-50.55,-2), -100)


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
