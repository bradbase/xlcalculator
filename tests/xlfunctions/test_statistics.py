import unittest

from xlcalculator.xlfunctions import xlerrors, func_xltypes, statistics


class StatisticsModuleTest(unittest.TestCase):

    def test_AVERAGE(self):
        self.assertEqual(
            statistics.AVERAGE(func_xltypes.Array([[1, 2], [3, 4]])), 2.5)

    def test_AVERAGE_numbers(self):
        self.assertEqual(statistics.AVERAGE(1, 2.0, 3, 4.0), 2.5)

    def test_AVERAGE_without_any_numbers(self):
        self.assertEqual(statistics.AVERAGE(), 0)

    def test_AVERAGE_mixed(self):
        self.assertEqual(
            statistics.AVERAGE(
                func_xltypes.Array([[1, 2], [3, 4]]), 1, 2, 3, 4
            ),
            2.5
        )

    def test_COUNT(self):
        range0 = func_xltypes.Array([[1, 2], [3, 4]])
        range1 = func_xltypes.Array([[1, 2], [3, 'SPAM']])
        self.assertEqual(statistics.COUNT(range0), 4)
        self.assertEqual(statistics.COUNT(range1), 3)
        self.assertEqual(statistics.COUNT(range0, range1), 7)
        self.assertEqual(statistics.COUNT(range0, range1, 1), 8)
        self.assertEqual(statistics.COUNT(range0, range1, 1, 'SPAM'), 8)

    def test_COUNT_without_any_values(self):
        self.assertIsInstance(
            statistics.COUNT(), xlerrors.ValueExcelError)

    def test_COUNT_with_too_many_values(self):
        self.assertIsInstance(
            statistics.COUNT([0] * 300), xlerrors.ValueExcelError)

    def test_COUNTA(self):
        range0 = func_xltypes.Array([[1, 2], [3, 4]])
        range1 = func_xltypes.Array([[2, 1], [3, '']])
        self.assertEqual(statistics.COUNTA(range0), 4)
        self.assertEqual(statistics.COUNTA(range1), 3)

    def test_COUNTA_with_bad_arg(self):
        self.assertIsInstance(statistics.COUNTA(None), xlerrors.NullExcelError)

    def test_COUNTA_without_any_values(self):
        self.assertIsInstance(statistics.COUNTA(), xlerrors.NullExcelError)

    def test_COUNTA_with_too_many_values(self):
        self.assertIsInstance(
            statistics.COUNTA([0] * 300), xlerrors.ValueExcelError)

    def test_COUNTIF(self):
        countRange = func_xltypes.Array([1, 2, 3, 4])
        condition = ">2"
        self.assertEqual(statistics.COUNTIF(countRange, condition), 2)

    def test_COUNTIF_text(self):
        countRange = func_xltypes.Array(["a", "b", "A"])
        condition = "=A"
        self.assertEqual(statistics.COUNTIF(countRange, condition), 2)

    def test_COUNTIFS(self):
        countRange = func_xltypes.Array([1, 2, 3, 4])
        countRange2 = func_xltypes.Array(["a", "B", "A", "A"])
        condition = "<3"
        condition2 = "A"
        self.assertEqual(statistics.COUNTIFS(
            countRange, condition, countRange2, condition2
        ), 1)

    def test_MAX(self):
        self.assertEqual(
            statistics.MAX(func_xltypes.Array([[1, 2], [3, 4]])),
            4
        )

    def test_MAX_without_any_numbers(self):
        self.assertEqual(statistics.MAX(), 0)

    def test_MIN(self):
        self.assertEqual(
            statistics.MIN(func_xltypes.Array([[1, 2], [3, 4]])),
            1
        )

    def test_MIN_without_any_numbers(self):
        self.assertEqual(statistics.MIN(), 0)

    def test_MIN_with_mixed_types(self):
        self.assertEqual(statistics.MIN(2, 3.0, True), 1)
