
import unittest

import pandas as pd

from koala_xlcalculator.function_library import xRound
from koala_xlcalculator.exceptions import ExcelError


class Test_Rounddown(unittest.TestCase):
    def test_nb_must_be_number(self):
        with self.assertRaises(ExcelError):
            xRound.rounddown('er', 1)


    def test_nb_digits_must_be_number(self):
        with self.assertRaises(ExcelError):
            xRound.rounddown(2.323, 'ze')


    def test_positive_number_of_digits(self):
        self.assertEqual(xRound.rounddown(-3.14159, 1), -3.1)


    def test_negative_number_of_digits(self):
        self.assertEqual(xRound.rounddown(31415.92654, -2), 31400)


    def test_round(self):
        self.assertEqual(xRound.rounddown(3.2, 0), 3)
        self.assertEqual(xRound.rounddown(76.9,0), 76)
        self.assertEqual(xRound.rounddown(3.14159, 3), 3.141)
