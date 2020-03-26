
import unittest

import pandas as pd

from koala_xlcalculator.function_library import Counta

"""
COUNTA(value1, [value2], ...)

The COUNTA function syntax has the following arguments:

value1    Required. The first argument representing the values that you want to count.

value2, ...    Optional. Additional arguments representing the values that you want to count, up to a maximum of 255 arguments.
"""

class TestCount(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_count(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        choose_result_00 = Counta.counta(range_00)
        result_00 = 4
        self.assertEqual(result_00, choose_result_00)

        range_01 = pd.DataFrame([[2, 1],[3, '']])
        choose_result_01 = Counta.counta(range_01)
        result_01 = 3
        self.assertEqual(result_01, choose_result_01)

        choose_result_02 = Counta.counta(range_00, range_01)
        result_02 = 7
        self.assertEqual(result_02, choose_result_02)

        choose_result_03 = Counta.counta(range_00, range_01, 1)
        result_03 = 8
        self.assertEqual(result_03, choose_result_03)

        choose_result_04 = Counta.counta(range_00, range_01, 1, '')
        result_04 = 8
        self.assertEqual(result_04, choose_result_04)

        choose_result_04 = Counta.counta(range_00, range_01, 1, None)
        result_04 = 8
        self.assertEqual(result_04, choose_result_04)


        with self.assertRaises(Exception) as context:
            Counta.counta(None)
            self.assertTrue('koala_xlcalculator.exceptions.exceptions.ExcelError: #VALUE' in context.exception)
