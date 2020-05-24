import unittest

from xlcalculator.types import XLCell


class CellTest(unittest.TestCase):

    def test_address(self):
        xlcell = XLCell('Sheet1!A1')
        address = 'Sheet1!A1'
        self.assertEqual(address, xlcell.address)

    def test_address_with_range(self):
        with self.assertRaises(Exception) as context:
            XLCell('Sheet1!A1:B1')
            self.assertTrue('This is a Range' in context.exception)

    def test_address_with_bad_sheet(self):
        # While the sheet name should be quoted, internally, the code often
        # just puts the sheet name in to produce unique keys, so the utility
        # supports unquoted sheets as well.
        self.assertEqual(XLCell('Bad Sheet!A1').address, 'Bad Sheet!A1')

    def test_value(self):
        cell = XLCell('Sheet1!A1', 5)
        self.assertEqual(cell.value, 5)

    def test_formula(self):
        cell = XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        self.assertEqual(cell.formula, 'SUM(A1:B1)')

    def test_float(self):
        cell = XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        self.assertEqual(float(cell), 5.0)

    def test_hash(self):
        cell = XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        self.assertEqual(hash(cell), hash(('Sheet1', 1, 1)))