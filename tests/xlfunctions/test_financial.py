import mock
import unittest

from xlcalculator.xlfunctions import (
    xl,
    xlerrors,
    func_xltypes,
    date,
    financial
)


class FinancialModuleTest(unittest.TestCase):

    def test_IRR(self):
        self.assertAlmostEqual(
            financial.IRR(
                func_xltypes.Array([[-100, 39, 59, 55, 20]])
            ),
            0.2809484
        )

    def test_IRR_with_guess_non_null(self):
        self.assertAlmostEqual(
            financial.IRR(func_xltypes.Array([[-100, 39, 59, 55, 20]]), 1),
            0.2809484)

    def test_NPV(self):
        self.assertAlmostEqual(
            financial.NPV(0.06, [[1, 2, 3]]), 5.2422470)
        self.assertAlmostEqual(
            financial.NPV(0.06, 1, 2, 3), 5.2422470)
        self.assertAlmostEqual(
            financial.NPV(0.06, 1), 0.9433962)
        self.assertAlmostEqual(
            financial.NPV(0.1, -10000, 3000, 4200, 6800), 1188.44, 2)

        range1 = func_xltypes.Array([[8000, 9200, 10000, 12000, 14500]])
        self.assertAlmostEqual(
            financial.NPV(0.08, range1) + -40000, 1922.06, 2)
        self.assertAlmostEqual(
            financial.NPV(0.08, range1, -9000) + -40000, -3749.47, 2)

    @mock.patch.object(xl, 'COMPATIBILITY', 'PYTHON')
    def test_NPV_in_python_mode(self):
        self.assertAlmostEqual(financial.NPV(0.06, 1, 2, 3), 5.556781772872908)
        self.assertAlmostEqual(financial.NPV(0.06, 1), 1.0)

    def test_NPV_without_values(self):
        self.assertIsInstance(financial.NPV(3.4), xlerrors.ValueExcelError)

    def test_PMT(self):
        self.assertAlmostEqual(
            financial.PMT(0.08 / 12, 10, 10000), -1037.03, 2
        )

    @mock.patch.object(xl, 'COMPATIBILITY', 'PYTHON')
    def test_PMT_in_python_mode(self):
        self.assertAlmostEqual(
            financial.PMT(0.08 / 12, 10, 10000), -1037.03, 2)
        self.assertAlmostEqual(
            financial.PMT(0.08 / 12, 10, 10000, type=1), -1030.164, 2)

    def test_PV(self):
        self.assertAlmostEqual(
            financial.PV(0.08 / 12, 12 * 20, 500.00, 0, 0),
            -59777.15,
            places=2
        )

    def test_SLN(self):
        self.assertEqual(financial.SLN(30000, 7500, 10), 2250)

    def test_VDB(self):
        self.assertAlmostEqual(
            financial.VDB(
                cost=575000,
                salvage=5000,
                life=10,
                start_period=3,
                end_period=5,
                factor=1.5,
            ),
            102160.546875
        )

    def test_VDB_partial(self):
        self.assertAlmostEqual(
            financial.VDB(
                cost=1,
                salvage=0,
                life=14,
                start_period=11.5,
                end_period=12.5,
                factor=1.25,
            ),
            0.068726290454684
        )

    def test_VDB_partial_no_switch(self):
        self.assertAlmostEqual(
            financial.VDB(
                cost=1,
                salvage=0,
                life=5.0,
                start_period=0.5,
                end_period=1.5,
                factor=2.5,
            ),
            0.375
        )

    def test_VDB_non_integer_life(self):
        self.assertAlmostEqual(
            financial.VDB(
                cost=575000,
                salvage=5000,
                life=10.5,
                start_period=3,
                end_period=5,
                factor=1.5,
            ),
            98708.55092429437
        )

    def test_XIRR(self):
        range_00 = func_xltypes.Array(
            [[-10000, 2750, 4250, 3250, 2750]])
        range_01 = func_xltypes.Array([[
            date.DATE(2008, 1, 1),
            date.DATE(2008, 3, 1),
            date.DATE(2008, 10, 30),
            date.DATE(2009, 2, 15),
            date.DATE(2009, 4, 1)
        ]])

        self.assertAlmostEqual(
            financial.XIRR(range_00, range_01, 0.1), 0.373362535)

    def test_XIRR_length_mismatch(self):
        range_00 = [[-10000, 2750, 4250]]
        range_01 = [[date.DATE(2008, 1, 1), date.DATE(2008, 3, 1)]]
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

    def test_XNPV(self):
        range_00 = func_xltypes.Array(
            [[-10000, 2750, 4250, 3250, 2750]])
        range_01 = func_xltypes.Array([[
            date.DATE(2008, 1, 1),
            date.DATE(2008, 3, 1),
            date.DATE(2008, 10, 30),
            date.DATE(2009, 2, 15),
            date.DATE(2009, 4, 1)
        ]])
        self.assertAlmostEqual(
            financial.XNPV(0.09, range_00, range_01), 2086.65, 2)

    def test_XNPV_length_mismatch(self):
        range_00 = [-10000, 2750, 4250]
        range_01 = [date.DATE(2008, 1, 1), date.DATE(2008, 3, 1)]
        self.assertIsInstance(
            financial.XNPV(0.09, range_00, range_01), xlerrors.NumExcelError)

    def test_XNPV_rate_lt_minus_one(self):
        range_00 = func_xltypes.Array(
            [[-10000, 2750, 4250, 3250, 2750]])
        range_01 = func_xltypes.Array([[
            date.DATE(2008, 1, 1),
            date.DATE(2008, 3, 1),
            date.DATE(2008, 10, 30),
            date.DATE(2009, 2, 15),
            date.DATE(2009, 4, 1)
        ]])
        self.assertEqual(
            financial.XNPV(-1.1, range_00, range_01), float('inf'))
