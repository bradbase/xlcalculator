
# Excel reference: https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Count
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator

"""
The COUNT function syntax has the following arguments:

value1    Required. The first item, cell reference, or range within which you want to count numbers.

value2, ...    Optional. Up to 255 additional items, cell references, or ranges within which you want to count numbers.

Note: The arguments can contain or refer to a variety of different types of data, but only numbers are counted.
"""

class TestCount(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/COUNT.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass

    def test_count(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        choose_result_00 = Count.count(range_00)
        result_00 = 4
        self.assertEqual(result_00, choose_result_00)

        range_01 = pd.DataFrame([[2, 1],[3, "SPAM"]])
        choose_result_01 = Count.count(range_01)
        result_01 = 3
        self.assertEqual(result_01, choose_result_01)

        choose_result_02 = Count.count(range_00, range_01)
        result_02 = 7
        self.assertEqual(result_02, choose_result_02)

        choose_result_03 = Count.count(range_00, range_01, 1)
        result_03 = 8
        self.assertEqual(result_03, choose_result_03)

        choose_result_04 = Count.count(range_00, range_01, 1, "SPAM")
        result_04 = 8
        self.assertEqual(result_04, choose_result_04)


    def test_count_evaluation_00(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_count_evaluation_02(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual( excel_value, value )


    @unittest.skip("There's a bug that doesn't create empty cells involved with formulas")
    def test_count_evaluation_03(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual( excel_value, value )
