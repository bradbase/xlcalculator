
# Excel reference: https://support.office.com/en-us/article/power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a

import unittest

from xlfunctions import Power
from xlfunctions.exceptions import ExcelError

from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from ..xlcalculator_test import XlCalculatorTestCase

class TestPower(XlCalculatorTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/POWER.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)
        

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqualTruncated( excel_value, value, 8 )


    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual( excel_value, value )
