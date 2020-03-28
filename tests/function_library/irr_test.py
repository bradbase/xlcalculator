
# Excel reference: https://support.office.com/en-us/article/irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc

import unittest

import pandas as pd

from koala_xlcalculator.function_library import IRR
from koala_xlcalculator.koala_types import XLCell
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestIRR(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/IRR.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass


    def test_irr_basic(self):
        range_00 = pd.DataFrame([[-100, 39, 59, 55, 20]])
        self.assertEqual(round(IRR.irr(range_00, 0), 7), 0.2809484)


    def test_irr_with_guess_non_null(self):
        with self.assertRaises(ValueError):
            IRR.irr([-100, 39, 59, 55, 20], 2)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( round(excel_value, 1), round(value, 1) )


    @unittest.skip("""Problem evalling: guess value for excellib.irr() is #N/A and not 0 for Sheet1!C1, IRR.irr(self.eval_ref("Sheet1!A2:A4"),Evaluator.apply_one("minus", 0.1, None, None))""")
    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual( excel_value, value )
