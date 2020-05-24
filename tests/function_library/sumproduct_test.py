
# Excel reference: https://support.office.com/en-us/article/sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e

import unittest

from xlcalculator.types import XLRange
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from . import testing


class TestSumProduct(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            testing.get_resource("SUMPRODUCT.xlsx"))
        self.evaluator = Evaluator(self.model)

    def test_evaluation_D7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D7')
        value = self.evaluator.evaluate('Sheet1!D7')
        self.assertEqual( excel_value, value )
