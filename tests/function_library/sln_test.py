
# Excel reference: https://support.office.com/en-us/article/sln-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8

import unittest

from xlcalculator.types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from . import testing


class TestSLN(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            testing.get_resource("SLN.xlsx"))
        self.evaluator = Evaluator(self.model)

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )
