
# Excel reference: https://support.office.com/en-us/article/power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a

import unittest

import pandas as pd

from koala_xlcalculator.function_library import Power
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator

from ..koala_test import KoalaTestCase

class TestPower(KoalaTestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/POWER.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    def test_first_argument_validity(self):
        self.assertEqual( 1, Power.power(-1, 2) )


    def test_second_argument_validity(self):
        self.assertEqual( 1, Power.power(1, 0) )


    def test_integers(self):
        self.assertEqual(Power.power(5, 2), 25)


    def test_floats(self):
        self.assertEqual(Power.power(98.6, 3.2), 2401077.2220695773)


    def test_fractions(self):
        self.assertEqual(Power.power(4,5/4), 5.656854249492381)


    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqualTruncated( excel_value, value, 8 )


    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual( excel_value, value )
