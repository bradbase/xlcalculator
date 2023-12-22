import unittest

from xlcalculator.xlfunctions import lookup, xlerrors, func_xltypes


class LookupModuleTest(unittest.TestCase):

    def test_CHOOSE(self):
        self.assertEqual(lookup.CHOOSE('2', 2, 4, 6), 4)

    def test_CHOOSE_with_negative_index(self):
        self.assertIsInstance(
            lookup.CHOOSE(-1, 1, 2, 3), xlerrors.ValueExcelError)

    def test_CHOOSE_with_too_large_index(self):
        self.assertIsInstance(
            lookup.CHOOSE(5, 1, 2, 3), xlerrors.ValueExcelError)

    def test_VLOOOKUP(self):
        # Excel Doc example.
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
            [102, 'Fortana', 'Olivier'],
            [103, 'Leal', 'Karina'],
            [104, 'Patten', 'Michael'],
            [105, 'Burke', 'Brian'],
            [106, 'Sousa', 'Luis'],
        ])
        self.assertEqual(lookup.VLOOKUP(102, range1, 2, False), 'Fortana')

    def test_VLOOOKUP_with_range_lookup(self):
        with self.assertRaises(NotImplementedError):
            lookup.VLOOKUP(1, func_xltypes.Array([[]]), 2, True)

    def test_VLOOOKUP_with_oversized_col_index_num(self):
        # Excel Doc example.
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
        ])
        self.assertIsInstance(
            lookup.VLOOKUP(102, range1, 4, False), xlerrors.ValueExcelError)

    def test_VLOOOKUP_with_unknown_lookup_value(self):
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
        ])
        self.assertIsInstance(
            lookup.VLOOKUP(102, range1, 2, False), xlerrors.NaExcelError)

    def test_VLOOOKUP_low_lookup_id(self):
        # Excel Doc example.
        range1 = func_xltypes.Array([
            [101, 'Davis', 'Sara'],
            [102, 'Fortana', 'Olivier'],
            [103, 'Leal', 'Karina'],
            [104, 'Patten', 'Michael'],
            [105, 'Burke', 'Brian'],
            [106, 'Sousa', 'Luis'],
        ])
        self.assertEqual(lookup.VLOOKUP(102, range1, 1, False), 'Fortana')

    def test_MATCH(self):
        range1 = [25, 28, 40, 41]
        self.assertEqual(lookup.MATCH(39, range1), 2)
        self.assertEqual(lookup.MATCH(39, range1, 1), 2)
        self.assertEqual(lookup.MATCH(40, range1, 1), 3)
        self.assertEqual(lookup.MATCH(40, range1, 0), 3)
        self.assertIsInstance(
            lookup.MATCH(40, range1, -1), xlerrors.NaExcelError
        )
        range2 = list(reversed(range1))
        self.assertEqual(lookup.MATCH(40, range2, -1), 2)
        self.assertIsInstance(
            lookup.MATCH(100, range2, -1), xlerrors.NaExcelError
        )
        self.assertIsInstance(
            lookup.MATCH(40, range2, 1), xlerrors.NaExcelError
        )
        self.assertIsInstance(
            lookup.MATCH(0, range1, 0), xlerrors.NaExcelError
        )
