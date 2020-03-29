
# Excel reference: https://support.microsoft.com/en-us/office/xnpv-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7

import unittest

import pandas as pd

from koala_xlcalculator.function_library import XNPV
from koala_xlcalculator.function_library import xDate
from koala_xlcalculator.koala_types import XLCell
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestNPV(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/XNPV.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass


    def test_npv_basic(self):
        range_00 = pd.DataFrame([[-10000, 2750, 4250, 3250, 2750]])
        range_01 = pd.DataFrame([[xDate.xdate(2008, 1, 1),
                                xDate.xdate(2008, 3, 1),
                                xDate.xdate(2008, 10, 30),
                                xDate.xdate(2009, 2, 15),
                                xDate.xdate(2009, 4, 1)
                                ]])
        self.assertEqual(round(XNPV.xnpv(0.09, range_00, range_01), 2), 2086.65)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )
