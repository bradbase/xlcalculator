import unittest

from xlcalculator.xlcalculator_types import XLCell


class TestCell(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_address(self):

        xlcell_00 = XLCell('Sheet1!A1')
        address_00 = 'Sheet1!A1'
        self.assertEqual(address_00, xlcell_00.address)

        with self.assertRaises(Exception) as context:
            XLCell('Sheet1!A1:B1')
            self.assertTrue('This is a Range' in context.exception)


    def test_value(self):

        xlcell_00 = XLCell('Sheet1!A1', 5)
        value_00 = 5
        self.assertEqual(value_00, xlcell_00.value)


    def test_formula(self):

        xlcell_00 = XLCell('Sheet1!A1', 5, 'SUM(A1:B1)')
        formula_00 = 'SUM(A1:B1)'
        self.assertEqual(formula_00, xlcell_00.formula)
