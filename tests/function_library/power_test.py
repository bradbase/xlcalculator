
import unittest

import pandas as pd

from koala_xlcalculator.function_library import Power
from koala_xlcalculator.exceptions import ExcelError


class Test_Power(unittest.TestCase):
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
