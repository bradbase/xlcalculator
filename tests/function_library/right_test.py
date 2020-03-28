# Excel reference: https://support.office.com/en-us/article/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Right
from koala_xlcalculator.koala_types import XLCell
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestNPV(unittest.TestCase):

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


    # @unittest.skip("""Problem evalling: unsupported operand type(s) for +: 'int' and 'XLCell' for Sheet1!A1, NPV.npv(self.eval_ref("Sheet1!A2"),self.eval_ref("Sheet1!A3"),self.eval_ref("Sheet1!A4"),self.eval_ref("Sheet1!A5"),self.eval_ref("Sheet1!A6"))""")
    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    # @unittest.skip("""Problem evalling: unsupported operand type(s) for +: 'int' and 'XLCell' for Sheet1!B1, Evaluator.apply("add",NPV.npv(self.eval_ref("Sheet1!B2"),self.eval_ref("Sheet1!B4:B8"),Evaluator.apply_one("minus", 9000, None, None)),self.eval_ref("Sheet1!B3"),None))""")
    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )
