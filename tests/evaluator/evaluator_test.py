
import unittest
import json
from copy import deepcopy

from jsonpickle import encode, decode

from koala_xlcalculator.evaluator.evaluator import Evaluator
from koala_xlcalculator.model.model import Model

from ..formulas import *


class TestEvaluator(unittest.TestCase):

    def setUp(self):
        self.model = Model()
        self.model.construct_from_json_file(r"./tests/resources/model.json")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)


    def test_evaluate(self):
        evaluated_result_00 = self.evaluator.evaluate('First!A2')
        result_00 = 0.1
        self.assertEqual(result_00, evaluated_result_00)

        evaluated_result_01 = self.evaluator.evaluate('Seventh!C1')
        result_01 = 3
        self.assertEqual(result_01, evaluated_result_01)

        evaluated_result_02 = self.evaluator.evaluate('Ninth!B1')
        result_02 = 101
        self.assertEqual(result_02, evaluated_result_02)

        evaluated_result_03 = self.evaluator.evaluate('Hundred')
        result_03 = 100
        self.assertEqual(result_03, evaluated_result_03)

        evaluated_result_04 = self.evaluator.evaluate('Tenth!C1')
        result_04 = 102
        self.assertEqual(result_04, evaluated_result_04)

        evaluated_result_05 = self.evaluator.evaluate('Tenth!C1')
        result_05 = 102
        self.assertEqual(result_05, evaluated_result_05)

        evaluated_result_06 = self.evaluator.evaluate('Tenth!C2')
        result_06 = 102
        self.assertEqual(result_06, evaluated_result_06)

        evaluated_result_07 = self.evaluator.evaluate('Tenth!C3')
        result_07 = 102
        self.assertEqual(result_07, evaluated_result_07)


    def test_set_value(self):
        self.evaluator.set_cell_value('First!A2', 88)
        evaluated_result_00 = self.evaluator.model.cells['First!A2'].value
        result_00 = 88
        self.assertEqual(result_00, evaluated_result_00)
        self.evaluator.set_cell_value('First!A2', 0.1) # Put it back the way we found it.


    def test_set_value_evaluate(self):
        self.evaluator.set_cell_value('First!A2', 88)
        evaluated_result_00 = self.evaluator.evaluate('First!A2')
        result_00 = 88
        self.assertEqual(result_00, evaluated_result_00)
        self.evaluator.set_cell_value('First!A2', 0.1) # Put it back the way we found it.
