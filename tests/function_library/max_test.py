import unittest

import pandas as pd

from koala_xlcalculator.function_library import xMax
from koala_xlcalculator.koala_types import XLRange


class TestMax(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_max(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        xsum_result_00 = xMax.xmax(range_00)
        result_00 = 4
        self.assertEqual(result_00, xsum_result_00)

        range_01 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        xsum_result_01 = xMax.xmax(range_01)
        result_01 = 4
        self.assertEqual(result_01, xsum_result_01)
