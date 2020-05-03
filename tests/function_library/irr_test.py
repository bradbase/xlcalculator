
# Excel reference: https://support.office.com/en-us/article/irr-function-64925eaa-9988-495b-b290-3ad0c163c1bc

import unittest

from xlfunctions import IRR
from xlfunctions.exceptions import ExcelError

from xlcalculator.xlcalculator_types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from ..xlcalculator_test import XlCalculatorTestCase


class TestIRR(XlCalculatorTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/IRR.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqualTruncated( excel_value, value )


    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqualTruncated( excel_value, value, 13 )


    @unittest.skip("""Problem evalling: guess value for excellib.irr() is #N/A and not 0 for Sheet1!C1, IRR.irr(self.eval_ref("Sheet1!A2:A4"),Evaluator.apply_one("minus", 0.1, None, None))""")
    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual( excel_value, value )
