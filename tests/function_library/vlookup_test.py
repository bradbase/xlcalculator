
# Excel reference: https://support.office.com/en-us/article/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

import unittest

import pandas as pd

from koala_xlcalculator.function_library import VLookup
from koala_xlcalculator.koala_types import XLCell
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestVLookup(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/VLOOKUP.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass


    @unittest.skip('Problem evalling: #VALUE! for Sheet1!A7, xRound.xround(Evaluator.apply_one("minus", 50.55, None, None))')
    def test_evaluation_B7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B7')
        value = self.evaluator.evaluate('Sheet1!B7')
        self.assertEqual( excel_value, value )
    #
    #
    # def test_evaluation_E7(self):
    #     excel_value = self.evaluator.get_cell_value('Sheet1!E7')
    #     value = self.evaluator.evaluate('Sheet1!E7')
    #     self.assertEqual( excel_value, value )
