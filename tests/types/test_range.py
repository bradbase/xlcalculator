import unittest

from xlcalculator.types import XLRange


class XLRangeTest(unittest.TestCase):

    def test_init(self):
        self.assertEqual(
            XLRange('Sheet1!A1:A4').address,
            [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3'], ['Sheet1!A4']]
        )
        self.assertEqual(
            XLRange('Sheet1!A1:A3,Sheet1!C1:C3,Sheet1!E1:E3').address,
            [['Sheet1!A1', 'Sheet1!C1', 'Sheet1!E1'],
             ['Sheet1!A2', 'Sheet1!C2', 'Sheet1!E2'],
             ['Sheet1!A3', 'Sheet1!C3', 'Sheet1!E3']]
        )

    def test_init_with_bad_sheet(self):
        # While the sheet name should be quoted, internally, the code often
        # just puts the sheet name in to produce unique keys, so the utility
        # supports unquoted sheets as well.
        self.assertEqual(
            XLRange('Bad Sheet!A1:A3').address,
            [['Bad Sheet!A1'], ['Bad Sheet!A2'], ['Bad Sheet!A3']]
        )

    def test_init_with_multiple_sheet(self):
        with self.assertRaises(ValueError):
            XLRange('Sheet1!A1:A4,Sheet2!A1:A4')

    def test_init_with_default_sheet(self):
        self.assertEqual(
            XLRange('A1:A3').address,
            [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3']]
        )

    def test_init_with_spaces(self):
        self.assertEqual(
            XLRange('Sheet1!A1:A3, Sheet1!C1:C3').address,
            [['Sheet1!A1', 'Sheet1!C1'],
             ['Sheet1!A2', 'Sheet1!C2'],
            ['Sheet1!A3', 'Sheet1!C3']]
        )

    def test_init_with_one_sheet_spec(self):
        self.assertEqual(
            XLRange('Sheet1!A1:A3,C1:C3').address,
            [['Sheet1!A1', 'Sheet1!C1'],
             ['Sheet1!A2', 'Sheet1!C2'],
             ['Sheet1!A3', 'Sheet1!C3']]
        )

    def test_cells(self):
        self.assertEqual(
            XLRange('Sheet1!A1:A3').cells,
            [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3']]
        )
