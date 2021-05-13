import unittest

from xlcalculator.xlfunctions import xlerrors, func_xltypes, engineering


class EngineeringModuleTest(unittest.TestCase):

    def test_DEC2BIN(self):
        self.assertEqual(
            engineering.DEC2BIN(10), func_xltypes.Text('1010'))

    def test_DEC2BIN_withPlaces(self):
        self.assertEqual(
            engineering.DEC2BIN(10, 5), func_xltypes.Text('01010'))

    def test_DEC2BIN_withBlank(self):
        self.assertEqual(
            engineering.DEC2BIN(func_xltypes.Blank()), func_xltypes.Text('0'))

    def test_DEC2BIN_withNegativeNumbers(self):
        self.assertEqual(engineering.DEC2BIN(-1), '1111111111')
        self.assertEqual(engineering.DEC2BIN(-2), '1111111110')
        self.assertEqual(engineering.DEC2BIN(-4.2), '1111111100')
        self.assertEqual(engineering.DEC2BIN(-4.8), '1111111100')

    def test_DEC2BIN_withBooleanNumber(self):
        self.assertIsInstance(
            engineering.DEC2BIN(True), xlerrors.ValueExcelError)

    def test_DEC2BIN_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.DEC2BIN(10, 20), xlerrors.NumExcelError)

    def test_DEC2BIN_withInsufficientPlaces(self):
        self.assertIsInstance(
            engineering.DEC2BIN(10, 3), xlerrors.NumExcelError)

    def test_DEC2BIN_withBooleanPlaces(self):
        self.assertIsInstance(
            engineering.DEC2BIN(10, True), xlerrors.ValueExcelError)

    def test_DEC2BIN_numberTooLarge(self):
        self.assertIsInstance(
            engineering.DEC2BIN(512), xlerrors.NumExcelError)
        self.assertIsInstance(
            engineering.DEC2BIN(-513), xlerrors.NumExcelError)

    def test_DEC2OCT(self):
        self.assertEqual(
            engineering.DEC2OCT(8), func_xltypes.Text('10'))
        self.assertEqual(
            engineering.DEC2OCT(-1), func_xltypes.Text('7777777777'))
        self.assertEqual(
            engineering.DEC2OCT(-2), func_xltypes.Text('7777777776'))

    def test_DEC2OCT_checkBounds(self):
        self.assertEqual(
            engineering.DEC2OCT(2**29 - 1), func_xltypes.Text('3777777777'))
        self.assertEqual(
            engineering.DEC2OCT(-(2**29)), func_xltypes.Text('4000000000'))
        self.assertIsInstance(
            engineering.DEC2OCT(2**29), xlerrors.NumExcelError)
        self.assertIsInstance(
            engineering.DEC2OCT(-(2**29) - 1), xlerrors.NumExcelError)

    def test_DEC2OCT_withPlaces(self):
        self.assertEqual(
            engineering.DEC2OCT(8, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.DEC2OCT(-1, 1), func_xltypes.Text('7777777777'))

    def test_DEC2OCT_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.DEC2OCT(8, 1), xlerrors.NumExcelError)

    def test_DEC2HEX(self):
        self.assertEqual(
            engineering.DEC2HEX(255), func_xltypes.Text('FF'))
        self.assertEqual(
            engineering.DEC2HEX(-1), func_xltypes.Text('FFFFFFFFFF'))
        self.assertEqual(
            engineering.DEC2HEX(-2), func_xltypes.Text('FFFFFFFFFE'))

    def test_DEC2HEX_checkBounds(self):
        self.assertEqual(
            engineering.DEC2HEX(2**39 - 1), func_xltypes.Text('7FFFFFFFFF'))
        self.assertEqual(
            engineering.DEC2HEX(-(2**39)), func_xltypes.Text('8000000000'))
        self.assertIsInstance(
            engineering.DEC2HEX(2**39), xlerrors.NumExcelError)
        self.assertIsInstance(
            engineering.DEC2HEX(-(2**39) - 1), xlerrors.NumExcelError)

    def test_DEC2HEX_withPlaces(self):
        self.assertEqual(
            engineering.DEC2HEX(16, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.DEC2HEX(-1, 1), func_xltypes.Text('FFFFFFFFFF'))

    def test_DEC2HEX_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.DEC2HEX(16, 1), xlerrors.NumExcelError)

    def test_BIN2DEC(self):
        self.assertEqual(
            engineering.BIN2DEC('1010'), 10)
        self.assertEqual(
            engineering.BIN2DEC(1010), 10)

    def test_BIN2DEC_withBlank(self):
        self.assertEqual(
            engineering.BIN2DEC(func_xltypes.Blank()), func_xltypes.Number(0))

    def test_BIN2DEC_withNonInteger(self):
        self.assertIsInstance(
            engineering.BIN2DEC(func_xltypes.Number(1.1)),
            xlerrors.NumExcelError)

    def test_BIN2DEC_numberTooLarge(self):
        self.assertIsInstance(
            engineering.BIN2DEC(10**20), xlerrors.NumExcelError)

    def test_BIN2DEC_withInvalidNumber(self):
        self.assertIsInstance(
            engineering.BIN2DEC('bad'), xlerrors.NumExcelError)

    def test_BIN2OCT(self):
        self.assertEqual(
            engineering.BIN2OCT(10000000), func_xltypes.Text('200'))
        self.assertEqual(
            engineering.BIN2OCT(1000000000), func_xltypes.Text('7777777000'))

    def test_BIN2OCT_checkBounds(self):
        self.assertEqual(
            engineering.BIN2OCT(1111111111), func_xltypes.Text('7777777777'))
        self.assertIsInstance(
            engineering.BIN2OCT(2), xlerrors.NumExcelError)
        self.assertIsInstance(
            engineering.BIN2OCT(10000000000), xlerrors.NumExcelError)

    def test_BIN2OCT_withPlaces(self):
        self.assertEqual(
            engineering.BIN2OCT(1000, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.BIN2OCT(1000000000, 1),
            func_xltypes.Text('7777777000'))

    def test_BIN2OCT_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.BIN2OCT(1000, 1), xlerrors.NumExcelError)

    def test_BIN2HEX(self):
        self.assertEqual(
            engineering.BIN2HEX(10), func_xltypes.Text('2'))
        self.assertEqual(
            engineering.BIN2HEX(10000000), func_xltypes.Text('80'))

    def test_BIN2HEX_checkBounds(self):
        self.assertEqual(
            engineering.BIN2HEX(1111111111), func_xltypes.Text('FFFFFFFFFF'))
        self.assertIsInstance(
            engineering.BIN2HEX(11000000000), xlerrors.NumExcelError)
        self.assertIsInstance(
            engineering.BIN2HEX(10000000000), xlerrors.NumExcelError)

    def test_BIN2HEX_withPlaces(self):
        self.assertEqual(
            engineering.BIN2HEX(1000, 5), func_xltypes.Text('00008'))
        self.assertEqual(
            engineering.BIN2HEX(1000000000, 1),
            func_xltypes.Text('FFFFFFFE00'))

    def test_BIN2HEX_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.BIN2HEX(10000, 1), xlerrors.NumExcelError)

    def test_OCT2DEC(self):
        self.assertEqual(
            engineering.OCT2DEC(10), 8)
        self.assertEqual(
            engineering.OCT2DEC(7777777777), -1)

    def test_OCT2DEC_checkBounds(self):
        self.assertEqual(
            engineering.OCT2DEC(3777777777), 536870911)
        self.assertEqual(
            engineering.OCT2DEC(4000000000), -536870912)
        self.assertIsInstance(
            engineering.OCT2DEC(10000000000), xlerrors.NumExcelError)

    def test_OCT2BIN(self):
        self.assertEqual(
            engineering.OCT2BIN(2), func_xltypes.Text('10'))
        self.assertEqual(
            engineering.OCT2BIN(777), func_xltypes.Text('111111111'))

    def test_OCT2BIN_checkBounds(self):
        self.assertEqual(
            engineering.OCT2BIN(7777777000), func_xltypes.Text('1000000000'))
        self.assertIsInstance(
            engineering.OCT2BIN(1000), xlerrors.NumExcelError)

    def test_OCT2BIN_withPlaces(self):
        self.assertEqual(
            engineering.OCT2BIN(2, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.OCT2BIN(7777777000, 1),
            func_xltypes.Text('1000000000'))

    def test_OCT2BIN_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.OCT2BIN(2, 1), xlerrors.NumExcelError)

    def test_OCT2HEX(self):
        self.assertEqual(
            engineering.OCT2HEX(10), func_xltypes.Text('8'))
        self.assertEqual(
            engineering.OCT2HEX(12), func_xltypes.Text('A'))

    def test_OCT2HEX_checkBounds(self):
        self.assertEqual(
            engineering.OCT2HEX(7777777777), func_xltypes.Text('FFFFFFFFFF'))
        self.assertIsInstance(
            engineering.OCT2HEX(10000000000), xlerrors.NumExcelError)

    def test_OCT2HEX_withPlaces(self):
        self.assertEqual(
            engineering.OCT2HEX(20, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.OCT2HEX(4000000000, 1),
            func_xltypes.Text('FFE0000000'))

    def test_OCT2HEX_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.OCT2HEX(20, 1), xlerrors.NumExcelError)

    def test_HEX2DEC(self):
        self.assertEqual(
            engineering.HEX2DEC(10), 16)
        self.assertEqual(
            engineering.HEX2DEC('fF'), 255)

    def test_HEX2DEC_checkBounds(self):
        self.assertEqual(
            engineering.HEX2DEC('FFFFFFFFFF'), -1)
        self.assertEqual(
            engineering.HEX2DEC(8000000000), -549755813888)
        self.assertIsInstance(
            engineering.HEX2DEC(10000000000), xlerrors.NumExcelError)

    def test_HEX2BIN(self):
        self.assertEqual(
            engineering.HEX2BIN(2), func_xltypes.Text('10'))
        self.assertEqual(
            engineering.HEX2BIN('1E0'), func_xltypes.Text('111100000'))

    def test_HEX2BIN_checkBounds(self):
        self.assertEqual(
            engineering.HEX2BIN('1FF'), func_xltypes.Text('111111111'))
        self.assertIsInstance(
            engineering.HEX2BIN('FFFFFFFDFF'), xlerrors.NumExcelError)

    def test_HEX2BIN_withPlaces(self):
        self.assertEqual(
            engineering.HEX2BIN(2, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.HEX2BIN('FFFFFFFE00', 1),
            func_xltypes.Text('1000000000'))

    def test_HEX2BIN_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.HEX2BIN(20, 1), xlerrors.NumExcelError)

    def test_HEX2OCT(self):
        self.assertEqual(
            engineering.HEX2OCT(8), func_xltypes.Text('10'))
        self.assertEqual(
            engineering.HEX2OCT('1E3'), func_xltypes.Text('743'))

    def test_HEX2OCT_checkBounds(self):
        self.assertEqual(
            engineering.HEX2OCT('1FFFFFFF'), func_xltypes.Text('3777777777'))
        self.assertEqual(
            engineering.HEX2OCT('FFE0000000'), func_xltypes.Text('4000000000'))
        self.assertIsInstance(
            engineering.HEX2OCT('FFDFFFFFFF'), xlerrors.NumExcelError)

    def test_HEX2OCT_withPlaces(self):
        self.assertEqual(
            engineering.HEX2OCT(8, 5), func_xltypes.Text('00010'))
        self.assertEqual(
            engineering.HEX2OCT('FFFFFFFE00', 1),
            func_xltypes.Text('7777777000'))

    def test_HEX2OCT_withTooManyPlaces(self):
        self.assertIsInstance(
            engineering.HEX2OCT(8, 1), xlerrors.NumExcelError)
