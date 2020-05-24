
# Excel reference: https://support.microsoft.com/en-us/office/xnpv-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7

import unittest

from xlcalculator.types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from . import testing

from ..xlcalculator_test import XlCalculatorTestCase


class TestNPV(XlCalculatorTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            testing.get_resource("XNPV.xlsx"))
        self.evaluator = Evaluator(self.model)

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqualTruncated( excel_value, value )
