
# Excel reference: https://support.office.com/en-us/article/npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568

import unittest

from xlfunctions import NPV
from xlfunctions.exceptions import ExcelError

from xlcalculator.xlcalculator_types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from ..xlcalculator_test import XlCalculatorTestCase


class TestNPV(XlCalculatorTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/NPV.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqualTruncated( excel_value, value, 11 )


    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )
