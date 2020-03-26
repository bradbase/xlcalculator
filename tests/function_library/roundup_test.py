
import unittest

import pandas as pd

from koala_xlcalculator.function_library import xRound
from koala_xlcalculator.exceptions import ExcelError


class Test_Roundup(unittest.TestCase):
    def test_nb_must_be_number(self):
        with self.assertRaises(ExcelError):
            xRound.roundup('er', 1)


    def test_nb_digits_must_be_number(self):
        with self.assertRaises(ExcelError):
            xRound.roundup(2.323, 'ze')


    def test_positive_number_of_digits(self):
        self.assertEqual(xRound.roundup(3.2,0), 4)


    def test_negative_number_of_digits(self):
        self.assertEqual(xRound.roundup(31415.92654, -2), 31500)


    def test_round(self):
        self.assertEqual(xRound.roundup(76.9,0), 77)
        self.assertEqual(xRound.roundup(3.14159, 3), 3.142)
        self.assertEqual(xRound.roundup(-3.14159, 1), -3.2)
