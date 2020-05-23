
# Excel reference: https://support.office.com/en-us/article/abs-function-3420200f-5628-4e8c-99da-c99d7c87713c

import unittest

from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from . import testing


class TestABS(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            testing.get_resource("ABS.xlsx"))
        self.evaluator = Evaluator(self.model)

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )
