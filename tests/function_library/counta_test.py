
# Excel reference: https://support.office.com/en-us/article/counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509

import unittest

import pandas as pd

from xlcalculator.function_library import Counta
from xlcalculator import ModelCompiler
from xlcalculator import Evaluator

"""
COUNTA(value1, [value2], ...)

The COUNTA function syntax has the following arguments:

value1    Required. The first argument representing the values that you want to count.

value2, ...    Optional. Additional arguments representing the values that you want to count, up to a maximum of 255 arguments.
"""

class TestCounta(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/COUNTA.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_counta(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        choose_result_00 = Counta.counta(range_00)
        result_00 = 4
        self.assertEqual(result_00, choose_result_00)

        range_01 = pd.DataFrame([[2, 1],[3, '']])
        choose_result_01 = Counta.counta(range_01)
        result_01 = 3
        self.assertEqual(result_01, choose_result_01)


        with self.assertRaises(Exception) as context:
            Counta.counta(None)
            self.assertTrue('xlcalculator.exceptions.exceptions.ExcelError: #VALUE' in context.exception)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( excel_value, value )


    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual( excel_value, value )
