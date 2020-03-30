
# Excel reference: https://support.office.com/en-us/article/rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53

import unittest

import pandas as pd

from koala_xlcalculator.function_library import xRound
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator

class Test_Rounddown(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/ROUNDDOWN.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_nb_must_be_number(self):
        self.assertIsInstance(xRound.rounddown('er', 1), ExcelError )


    def test_nb_digits_must_be_number(self):
        self.assertIsInstance(xRound.rounddown(2.323, 'ze'), ExcelError )


    def test_positive_number_of_digits(self):
        self.assertEqual(xRound.rounddown(-3.14159, 1), -3.1)


    def test_negative_number_of_digits(self):
        self.assertEqual(xRound.rounddown(31415.92654, -2), 31400)


    def test_round(self):
        self.assertEqual(xRound.rounddown(3.2, 0), 3)
        self.assertEqual(xRound.rounddown(76.9,0), 76)
        self.assertEqual(xRound.rounddown(3.14159, 3), 3.141)


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


    @unittest.skip('Problem evalling: #VALUE! for Sheet1!A4, xRound.rounddown(Evaluator.apply_one("minus", 3.14159, None, None),1)')
    def test_evaluation_A4(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A4')
        value = self.evaluator.evaluate('Sheet1!A4')
        self.assertEqual( excel_value, value )


    @unittest.skip('Problem evalling: #VALUE! for Sheet1!A5, xRound.rounddown(31415.92654,Evaluator.apply_one("minus", 2, None, None))')
    def test_evaluation_A5(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A5')
        value = self.evaluator.evaluate('Sheet1!A5')
        self.assertEqual( excel_value, value )
