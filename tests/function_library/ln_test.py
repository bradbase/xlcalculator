
# Excel reference: https://support.office.com/en-us/article/ln-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f

import unittest

import pandas as pd

from xlcalculator.function_library import Ln
from xlcalculator.xlcalculator_types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from ..xlcalculator_test import XlCalculatorTestCase


class TestLn(XlCalculatorTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/LN.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    def test_ln(self):
        ln_result_00 = Ln.ln(86)
        result_00 = 4.4543473
        self.assertEqual(result_00, round(ln_result_00, 7))


    def test_ln_rounded(self):
        ln_01 = XLCell("Sheet1!A1", 86)
        result_01 = 4.4543473 # this doesn't have enough decimal places but is what's in the LN function examples from MS.
        ln_result_01 = Ln.ln(ln_01)
        self.assertEqualRounded(result_01, ln_result_01)


    def test_ln_not_rounded(self):
        ln_result_02 = Ln.ln(2.7182818)
        result_02 = 1 # this is in error but is what's in the LN function examples from MS.
        self.assertEqualRounded(result_02, ln_result_02)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqualTruncated( excel_value, value )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqualTruncated( excel_value, value )
