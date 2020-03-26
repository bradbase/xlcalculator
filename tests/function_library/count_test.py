
import unittest

import pandas as pd

from koala_xlcalculator.function_library import Count

"""
The COUNT function syntax has the following arguments:

value1    Required. The first item, cell reference, or range within which you want to count numbers.

value2, ...    Optional. Up to 255 additional items, cell references, or ranges within which you want to count numbers.

Note: The arguments can contain or refer to a variety of different types of data, but only numbers are counted.
"""

class TestCount(unittest.TestCase):

    # def setUp(self):
    #     pass

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

        range_01 = pd.DataFrame([[2, 1],[3, "SPAM"]])
        choose_result_01 = Count.count(range_00, range_01)
        result_01 = 7
        self.assertEqual(result_01, choose_result_01)
