import unittest

from xlcalculator import xltypes
from xlcalculator.tokenizer import f_token

from . import testing


class XLFormulaTest(testing.XlCalculatorTestCase):

    def test_formula(self):
        self.assertEqual(
            xltypes.XLFormula('=SUM(A1:B1)').formula, '=SUM(A1:B1)')

    def test_tokens(self):
        self.assertASTNodesEqual(
            xltypes.XLFormula('=SUM(A1:B1)').tokens,
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop')
            ]
        )

    def test_defined_name_formula(self):
        self.assertASTNodesEqual(
            xltypes.XLFormula('=SUM(A1:B1)').tokens,
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop')
            ]
        )


class XLCellTest(unittest.TestCase):

    def test_address(self):
        xlcell = xltypes.XLCell('Sheet1!A1')
        address = 'Sheet1!A1'
        self.assertEqual(address, xlcell.address)

    def test_address_with_range(self):
        with self.assertRaises(Exception) as context:
            xltypes.XLCell('Sheet1!A1:B1')
            self.assertTrue('This is a Range' in context.exception)

    def test_address_with_bad_sheet(self):
        # While the sheet name should be quoted, internally, the code often
        # just puts the sheet name in to produce unique keys, so the utility
        # supports unquoted sheets as well.
        self.assertEqual(
            xltypes.XLCell('Bad Sheet!A1').address, 'Bad Sheet!A1')

    def test_value(self):
        cell = xltypes.XLCell('Sheet1!A1', 5)
        self.assertEqual(cell.value, 5)

    def test_formula(self):
        cell = xltypes.XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        self.assertEqual(cell.formula, 'SUM(A1:B1)')

    def test_float(self):
        cell = xltypes.XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        self.assertEqual(float(cell), 5.0)

    def test_hash(self):
        cell = xltypes.XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        self.assertEqual(hash(cell), hash(('Sheet1', 1, 1)))


class XLRangeTest(unittest.TestCase):

    def test_init(self):
        self.assertEqual(
            xltypes.XLRange('Sheet1!A1:A4').address,
            [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3'], ['Sheet1!A4']]
        )
        self.assertEqual(
            xltypes.XLRange('Sheet1!A1:A3,Sheet1!C1:C3,Sheet1!E1:E3').address,
            [['Sheet1!A1', 'Sheet1!C1', 'Sheet1!E1'],
             ['Sheet1!A2', 'Sheet1!C2', 'Sheet1!E2'],
             ['Sheet1!A3', 'Sheet1!C3', 'Sheet1!E3']]
        )

    def test_init_with_bad_sheet(self):
        # While the sheet name should be quoted, internally, the code often
        # just puts the sheet name in to produce unique keys, so the utility
        # supports unquoted sheets as well.
        self.assertEqual(
            xltypes.XLRange('Bad Sheet!A1:A3').address,
            [['Bad Sheet!A1'], ['Bad Sheet!A2'], ['Bad Sheet!A3']]
        )

    def test_init_with_multiple_sheet(self):
        with self.assertRaises(ValueError):
            xltypes.XLRange('Sheet1!A1:A4,Sheet2!A1:A4')

    def test_init_with_default_sheet(self):
        self.assertEqual(
            xltypes.XLRange('A1:A3').address,
            [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3']]
        )

    def test_init_with_spaces(self):
        self.assertEqual(
            xltypes.XLRange('Sheet1!A1:A3, Sheet1!C1:C3').address,
            [['Sheet1!A1', 'Sheet1!C1'],
             ['Sheet1!A2', 'Sheet1!C2'],
             ['Sheet1!A3', 'Sheet1!C3']]
        )

    def test_init_with_one_sheet_spec(self):
        self.assertEqual(
            xltypes.XLRange('Sheet1!A1:A3,C1:C3').address,
            [['Sheet1!A1', 'Sheet1!C1'],
             ['Sheet1!A2', 'Sheet1!C2'],
             ['Sheet1!A3', 'Sheet1!C3']]
        )

    def test_cells(self):
        self.assertEqual(
            xltypes.XLRange('Sheet1!A1:A3').cells,
            [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3']]
        )
