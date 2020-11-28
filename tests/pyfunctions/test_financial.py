
import unittest
from unittest import mock

from xlcalculator.xlfunctions import (
    xlerrors,
    func_xltypes,
    date
)

from xlcalculator.pyfunctions import (
    financial
)


class PyFunctions_FinancialModuleTest(unittest.TestCase):

    def test_xirr(self):
        dates = [float(date.DATE(2010, 12, 29)),
                 float(date.DATE(2012, 1, 25)),
                 float(date.DATE(2012, 3, 8))]
        values = [-10000, 20, 10100]
        self.assertEqual(
            financial._xirr(values, dates, -0.99), 0.01006126516492058)

    def test_xirr_singleEntry(self):
        dates = [float(date.DATE(2019, 1, 1))]
        values = [-10000]
        self.assertEqual(
            financial._xirr(values, dates, -0.99), 0)

    def test_xirr_noConvergence(self):
        dates = [float(date.DATE(2019, 1, 1)), float(date.DATE(2020, 1, 1))]
        values = [-10000, -10000]
        self.assertEqual(
            financial._xirr(values, dates, -0.99), -0.9999999999989412)

    def test_XIRR_xirr(self):
        range_00 = func_xltypes.Array(
            [[-10000, 20, 10100]])
        range_01 = func_xltypes.Array([[
            date.DATE(2010, 12, 29),
            date.DATE(2012, 1, 25),
            date.DATE(2012, 3, 8)
        ]])

        dates = [float(date.DATE(2010, 12, 29)),
                 float(date.DATE(2012, 1, 25)),
                 float(date.DATE(2012, 3, 8))]
        values = [-10000, 20, 10100]

        target = func_xltypes.XlNumber(financial._xirr(values, dates, 0.1))
        test = financial.XIRR(range_00, range_01, 0.1)
        self.assertEqual(test, target)

    def test_XIRR(self):
        range_00 = func_xltypes.Array(
            [[-10000, 20, 10100]])
        range_01 = func_xltypes.Array([[
            date.DATE(2010, 12, 29),
            date.DATE(2012, 1, 25),
            date.DATE(2012, 3, 8)
        ]])

        self.assertEqual(
            financial.XIRR(range_00, range_01, -0.99), 0.01006126516492058)

    def test_XIRR_allPositiveValues(self):
        range_00 = func_xltypes.Array(
            [[10000, 20, 10100]])
        range_01 = func_xltypes.Array([[
            date.DATE(2010, 12, 29),
            date.DATE(2012, 1, 25),
            date.DATE(2012, 3, 8)
        ]])

        self.assertIsInstance(financial.XIRR(range_00, range_01, 0.1),
                              xlerrors.NumExcelError)

    def test_XIRR_allNegativeValues(self):
        range_00 = func_xltypes.Array(
            [[-10000, -20, -10100]])
        range_01 = func_xltypes.Array([[
            date.DATE(2010, 12, 29),
            date.DATE(2012, 1, 25),
            date.DATE(2012, 3, 8)
        ]])

        self.assertIsInstance(financial.XIRR(range_00, range_01, 0.1),
                              xlerrors.NumExcelError)

    def test_XIRR_withRuntimeError(self):
        range_00 = func_xltypes.Array(
            [[-10000, 20, 10100]])
        range_01 = func_xltypes.Array([[
            date.DATE(2010, 12, 29),
            date.DATE(2012, 1, 25),
            date.DATE(2012, 3, 8)
        ]])

        with mock.patch.object(financial, '_xirr',
                               side_effect=RuntimeError()):
            self.assertIsInstance(financial.XIRR(range_00, range_01, 0.1),
                                  xlerrors.NumExcelError)

    def test_XIRR_withValueError(self):
        range_00 = func_xltypes.Array(
            [[-10000, 20, 10100]])
        range_01 = func_xltypes.Array([[
            date.DATE(2010, 12, 29),
            date.DATE(2012, 1, 25),
            date.DATE(2012, 3, 8)
        ]])

        with mock.patch.object(financial, '_xirr',
                               side_effect=ValueError()):
            self.assertIsInstance(financial.XIRR(range_00, range_01, 0.1),
                                  xlerrors.NumExcelError)

    def test_XIRR_dates_length_mismatch(self):
        range_00 = [[-10000, 2750, 4250]]
        range_01 = [[date.DATE(2008, 1, 1), date.DATE(2008, 3, 1)]]

        self.assertIsInstance(
            financial.XIRR(range_00, range_01, 0.1), xlerrors.NumExcelError)

    def test_XIRR_values_length_mismatch(self):
        range_00 = [[-10000, 2750]]
        range_01 = [[date.DATE(2008, 1, 1), date.DATE(2008, 3, 1), date.DATE(2008, 10, 30)]]

        self.assertIsInstance(
            financial.XIRR(range_00, range_01, 0.1), xlerrors.NumExcelError)

    def test_XIRR_not_converge(self):
        dates = [date.DATE(2020, 8, 31),
                 date.DATE(2020, 5, 5),
                 date.DATE(2020, 2, 28),
                 date.DATE(2020, 8, 31),
                 date.DATE(2018, 6, 30)]
        values = [50289.0, -75000.0, 0.0, 0.0, 0.0]
        self.assertIsInstance(
            financial.XIRR(values, dates, 0), xlerrors.NumExcelError)
