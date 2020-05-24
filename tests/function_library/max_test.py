
# Excel reference: https://support.office.com/en-us/article/max-function-e0012414-9ac8-4b34-9a47-73e662c08098

import unittest

from xlcalculator.types import XLRange
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from . import testing


class TestMax(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            testing.get_resource("MAX.xlsx"))
        self.evaluator = Evaluator(self.model)

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )

    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )
