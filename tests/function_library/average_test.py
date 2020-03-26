import unittest
import json
from copy import deepcopy

import pandas as pd
from jsonpickle import encode, decode

from koala_xlcalculator.function_library import Average
from koala_xlcalculator.koala_types import XLRange
from koala_xlcalculator.evaluator.evaluator import Evaluator
from koala_xlcalculator.model.model import Model


class TestAverage(unittest.TestCase):

    def setUp(self):
        self.model = Model()
        self.model.construct_from_json_file(r"./tests/resources/average.json")
        self.evaluator = Evaluator(self.model)

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

        range_02 = XLRange("Sheet1!A1:B2", "Sheet1!A1:B2", value = pd.DataFrame([[1, 2],[3, 4]]))
        average_result_02 = Average.average(range_01, 5)
        result_02 = 3
        self.assertEqual(result_02, average_result_02)


    def test_average_evaluation(self):
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( 2.5, value )
