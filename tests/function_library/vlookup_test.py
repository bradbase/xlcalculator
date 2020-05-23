
# Excel reference: https://support.office.com/en-us/article/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

import unittest

from xlcalculator.xlcalculator_types import XLCell
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

from . import testing


class TestVLookup(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(
            testing.get_resource("VLOOKUP.xlsx"))
        self.evaluator = Evaluator(self.model)

    @unittest.skip("""Problem evalling: Excact match only supported at the moment. for Sheet1!B7, VLookup.vlookup(eval_ref("Sheet1!A7"),eval_ref("Sheet1!A2:B5"),2,True)""")
    def test_evaluation_B7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B7')
        value = self.evaluator.evaluate('Sheet1!B7')
        self.assertEqual( excel_value, value )


    def test_evaluation_E7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!E7')
        value = self.evaluator.evaluate('Sheet1!E7')
        self.assertEqual( excel_value, value )
