import unittest

import pandas as pd

from koala_xlcalculator.function_library import xMin
from koala_xlcalculator.koala_types import XLRange


class TestMin(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_min(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        xsum_result_00 = xMin.xmin(range_00)
        result_00 = 1
        self.assertEqual(result_00, xsum_result_00)

        range_01 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        xsum_result_01 = xMin.xmin(range_01)
        result_01 = 1
        self.assertEqual(result_01, xsum_result_01)
