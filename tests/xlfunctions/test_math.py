import unittest

from xlcalculator.xlfunctions import math, xlerrors, func_xltypes


class DateModuleTest(unittest.TestCase):

    def test_ABS(self):
        self.assertEqual(math.ABS(3), 3)
        self.assertEqual(math.ABS('-3.4'), 3.4)

    def test_ABS_with_bad_arg(self):
        self.assertIsInstance(math.ABS('bad'), xlerrors.ValueExcelError)

    def test_LN(self):
        self.assertEqual(math.LN(2.718281828459045), 1)

    def test_LN_with_bad_arg(self):
        self.assertIsInstance(math.LN('bad'), xlerrors.ValueExcelError)

    def test_MOD(self):
        self.assertEqual(math.MOD(1, 2), 1)

    def test_MOD_with_float(self):
        self.assertEqual(math.MOD(1.0, 2), 1)

    def test_MOD_with_string(self):
        self.assertEqual(math.MOD('1.0', 2), 1)

    def test_MOD_with_bad_arg(self):
        self.assertIsInstance(math.MOD('bad', 2), xlerrors.ValueExcelError)
        self.assertIsInstance(math.MOD(1, 'bad'), xlerrors.ValueExcelError)

    def test_PI(self):
        self.assertEqual(math.PI(), 3.141592653589793)

    def test_POWER(self):
        self.assertEqual(math.POWER(10, 2), 100)
        self.assertEqual(math.POWER('-1', 2), 1)
        self.assertEqual(math.POWER(2.5, 2), 6.25)

    def test_POWER_with_bad_arg(self):
        self.assertIsInstance(math.POWER('bad', 2), xlerrors.ValueExcelError)
        self.assertIsInstance(math.POWER(10, 'bad'), xlerrors.ValueExcelError)

    def test_ROUND(self):
        self.assertEqual(math.ROUND(0.6), 1)
        self.assertEqual(math.ROUND(1.3), 1)
        self.assertEqual(math.ROUND(1.25, 1), 1.3)

    def test_ROUND_with_bad_arg(self):
        self.assertIsInstance(
            math.ROUND('bad'), xlerrors.ValueExcelError)
        self.assertIsInstance(
            math.ROUND(1.3, 'bad'), xlerrors.ValueExcelError)

    def test_ROUNDUP(self):
        self.assertEqual(math.ROUNDUP(0.6), 1)
        self.assertEqual(math.ROUNDUP(1.3), 2)
        self.assertEqual(math.ROUNDUP(1.24, 1), 1.3)

    def test_ROUNDUP_with_bad_arg(self):
        self.assertIsInstance(
            math.ROUNDUP('bad'), xlerrors.ValueExcelError)
        self.assertIsInstance(
            math.ROUNDUP(1.3, 'bad'), xlerrors.ValueExcelError)

    def test_ROUNDDOWN(self):
        self.assertEqual(math.ROUNDDOWN(0.6), 0)
        self.assertEqual(math.ROUNDDOWN(1.3), 1)
        self.assertEqual(math.ROUNDDOWN(1.26, 1), 1.2)

    def test_ROUNDDOWN_with_bad_arg(self):
        self.assertIsInstance(
            math.ROUNDDOWN('bad'), xlerrors.ValueExcelError)
        self.assertIsInstance(
            math.ROUNDDOWN(1.3, 'bad'), xlerrors.ValueExcelError)

    def test_SQRT(self):
        self.assertEqual(math.SQRT(4), 2)
        self.assertEqual(math.SQRT(4.0), 2.0)

    def test_SQRT_with_neg_number(self):
        self.assertIsInstance(math.SQRT(-4), xlerrors.NumExcelError)

    def test_SQRT_with_bad_arg(self):
        self.assertIsInstance(math.SQRT('bad'), xlerrors.ValueExcelError)

    def test_SUM(self):
        self.assertEqual(math.SUM(func_xltypes.Array([[1, 2], [3, 4]])), 10)
        self.assertEqual(math.SUM(1, 2, 3, 4.0), 10.0)

    def test_SUM_with_nonnumbers_in_range(self):
        self.assertEqual(math.SUM(func_xltypes.Array([[1, 'bad'], [3, 4]])), 8)
        self.assertEqual(math.SUM(
            func_xltypes.Array([
                [func_xltypes.Number(1), func_xltypes.Text('N/A')],
                [func_xltypes.Number(3), func_xltypes.Number(4)]
            ])), 8)

    def test_SUM_with_bad_Arg(self):
        self.assertEqual(math.SUM('foo'), 0)

    def test_SUM_empty(self):
        self.assertEqual(math.SUM(), 0)

    def test_SUMIF(self):
        self.assertEqual(math.SUMIF([0, 1, 2], '>=1', [10, 20, 30]), 50)

    def test_SUMIF_invalid_criteria(self):
        self.assertIsInstance(
            math.SUMIF([0, 1, 2], [0, 1], [10, 20, 30]),
            xlerrors.ValueExcelError)

    def test_SUMIF_unspecified_sum_range(self):
        self.assertEqual(math.SUMIF([0, 1, 2, 3], ">=2"), 5)

    def test_SUMIF_with_invalid_sum_range(self):
        # In this case, "bad" is converted to a single item array, then
        # filtered to an array where the value is 0, so that the sum is always
        # 0.
        self.assertEqual(math.SUMIF([0, 1, 2, 3], ">=2", 'bad'), 0)

    def test_SUMIFS(self):
        self.assertEqual(
            math.SUMIFS([10, 20, 30], [0, 1, 2], ">=1", ["a", "b", "a"], "a"),
            30
        )

    def test_SUMIFS_invalid_criteria(self):
        self.assertIsInstance(
            math.SUMIFS([10, 20, 30], [0, 1, 2], [0, 1], ["a", "b", "a"], "a"),
            xlerrors.ValueExcelError
        )

    def test_SUMIFS_unequal_array_lengths(self):
        self.assertEqual(
            math.SUMIFS(
                [10, 20, 30], [0, 1, 2], ">=1", ["a", "b", "a", 1], "a"
            ),
            0
        )

    def test_SUMPRODUCT(self):
        range1 = func_xltypes.Array([[1], [10], [3]])
        range2 = func_xltypes.Array([[3], [1], [2]])
        self.assertEqual(math.SUMPRODUCT(range1, range2), 19)

    def test_SUMPRODUCT_ranges_with_different_sizes(self):
        range1 = func_xltypes.Array([[1], [10], [3]])
        range2 = func_xltypes.Array([[3], [3], [1], [2]])
        self.assertIsInstance(
            math.SUMPRODUCT(range1, range2), xlerrors.ValueExcelError)

    def test_SUMPRODUCT_with_empty_frist_range(self):
        self.assertEqual(math.SUMPRODUCT(func_xltypes.Array([])), 0)

    def test_SUMPRODUCT_without_any_range(self):
        self.assertIsInstance(math.SUMPRODUCT(), xlerrors.NullExcelError)

    def test_SUMPRODUCT_ranges_with_errors(self):
        range1 = func_xltypes.Array(
            [[xlerrors.NumExcelError('err')], [10], [3]]
        )
        range2 = func_xltypes.Array([[3], [3], [1]])
        self.assertIsInstance(
            math.SUMPRODUCT(range1, range2), xlerrors.NaExcelError)

    def test_SUMPRODUCT_with_single_value(self):
        self.assertEqual(math.SUMPRODUCT(1), 1.0)

    def test_TRUNC(self):
        self.assertEqual(math.TRUNC(0.6), 0)
        self.assertEqual(math.TRUNC(1.3), 1)
        self.assertEqual(math.TRUNC(1.26, 1), 1.2)

    def test_TRUNC_with_bad_arg(self):
        self.assertIsInstance(
            math.TRUNC('bad'), xlerrors.ValueExcelError)
        self.assertIsInstance(
            math.TRUNC(1.3, 'bad'), xlerrors.ValueExcelError)
