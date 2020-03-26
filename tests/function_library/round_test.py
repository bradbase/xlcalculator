
import unittest

import pandas as pd

from koala_xlcalculator.function_library import xRound
from koala_xlcalculator.exceptions import ExcelError


class Test_Round(unittest.TestCase):
    def test_nb_must_be_number(self):
        with self.assertRaises(ExcelError):
            xRound.xround('er', 1)


    def test_nb_digits_must_be_number(self):
        with self.assertRaises(ExcelError):
            xRound.xround(2.323, 'ze')


    def test_positive_number_of_digits(self):
        self.assertEqual(xRound.xround(2.675, 2), 2.68)


    def test_negative_number_of_digits(self):
        self.assertEqual(xRound.xround(2352.67, -2), 2400)

    def test_round(self):
        self.assertEqual(xRound.xround(2.15, 1), 2.2)
        self.assertEqual(xRound.xround(2.149, 1), 2.1)
        self.assertEqual(xRound.xround(-1.475, 2), -1.48)
        self.assertEqual(xRound.xround(21.5, -1), 20)
        self.assertEqual(xRound.xround(626.3,-3), 1000)
        self.assertEqual(xRound.xround(1.98,-1), 0)
        self.assertEqual(xRound.xround(-50.55,-2), -100)
