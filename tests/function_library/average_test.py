import unittest

import pandas as pd

from koala_xlcalculator.function_library import Average
from koala_xlcalculator.koala_types import XLRange


class TestAverage(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_average(self):
        range_00 = pd.DataFrame([[1, 2],[3, 4]])
        average_result_00 = Average.average(range_00)
        result_00 = 2.5
        self.assertEqual(result_00, average_result_00)

        range_01 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        average_result_01 = Average.average(range_01)
        result_01 = 2.5
        self.assertEqual(result_01, average_result_01)
