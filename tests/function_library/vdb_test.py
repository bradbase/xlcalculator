# Excel reference: https://support.office.com/en-us/article/vdb-function-dde4e207-f3fa-488d-91d2-66d55e861d73

import unittest

import pandas as pd

from koala_xlcalculator.function_library import VDB
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class Test_VDB(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/VDB.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    def test_vdb_basic(self):
        cost = 575000
        salvage = 5000
        life = 10
        rate = 1.5
        start = 3
        end = 5

        obj = 102160.546875

        self.assertEqual(VDB.vdb(cost, salvage, life, start, end, rate), obj)

    def test_vdb_partial(self):
        cost = 1
        salvage = 0
        life = 14
        rate = 1.25
        start = 11.5
        end = 12.5

        obj = 0.068726290454684

        self.assertEqual(round(VDB.vdb(cost, salvage, life, start, end, rate), 15), obj)

    def test_vdb_partial_no_switch(self):
        cost = 1
        salvage = 0
        life = 5.0
        rate = 2.5
        start = 0.5
        end = 1.5

        obj = 0.375

        self.assertEqual(VDB.vdb(cost, salvage, life, start, end, rate, True), obj)


    def test_evaluation_F23(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!F23')
        value = self.evaluator.evaluate('Sheet1!F23')
        self.assertEqual( excel_value, value )


    # def test_evaluation_F24(self):
    #     excel_value = self.evaluator.get_cell_value('Sheet1!F24')
    #     value = self.evaluator.evaluate('Sheet1!F24')
    #     self.assertEqual( excel_value, value )
    #
    #
    # def test_evaluation_F27(self):
    #     excel_value = self.evaluator.get_cell_value('Sheet1!F27')
    #     value = self.evaluator.evaluate('Sheet1!F27')
    #     self.assertEqual( excel_value, value )
    #
    #
    # def test_evaluation_F28(self):
    #     excel_value = self.evaluator.get_cell_value('Sheet1!F28')
    #     value = self.evaluator.evaluate('Sheet1!F28')
    #     self.assertEqual( excel_value, value )
