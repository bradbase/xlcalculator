
# Excel reference: https://support.office.com/en-us/article/sln-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8

import unittest

import pandas as pd

from koala_xlcalculator.function_library import SLN
from koala_xlcalculator.koala_types import XLCell
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestSLN(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/SLN.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_sln(self):
        sln_result_00 = SLN.sln(30000, 7500, 10)
        result_00 = 2250
        self.assertEqual(result_00, sln_result_00)

        cell_01 = XLCell("Sheet1!A2", 30000)
        cell_02 = XLCell("Sheet1!A3", 7500)
        cell_03 = XLCell("Sheet1!A4", 10)
        xsum_result_01 = SLN.sln(cell_01, cell_02, cell_03)
        result_01 = 2250
        self.assertEqual(result_01, xsum_result_01)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )
