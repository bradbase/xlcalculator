# Excel reference: https://support.office.com/en-us/article/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f

import unittest

import pandas as pd

from xlcalculator.function_library import Right
from xlcalculator.xlcalculator_types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator


class TestRight(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/RIGHT.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass


    def test_right(self):
        self.assertEqual(Right.right("Sale Price", 5), "Price")
        self.assertEqual(Right.right("Stock Number"), "r")


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )
