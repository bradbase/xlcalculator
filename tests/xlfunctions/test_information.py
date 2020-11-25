import unittest

from xlcalculator.xlfunctions import information, xlerrors, func_xltypes, date


class InformationModuleTest(unittest.TestCase):

    def test_ISBLANK(self):
        self.assertTrue(information.ISBLANK(func_xltypes.Blank()))
        self.assertFalse(information.ISBLANK(func_xltypes.Text("hello")))
        self.assertTrue(information.ISBLANK(func_xltypes.Text("")))

    def test_ISTEXT(self):
        self.assertTrue(information.ISTEXT(func_xltypes.Text("hello")))
        self.assertFalse(information.ISTEXT(func_xltypes.Number(1234)))
        self.assertTrue(information.ISTEXT("hello"))
        self.assertFalse(information.ISTEXT(1234))
        self.assertTrue(information.ISTEXT("1234"))

    def test_NA(self):
        self.assertEqual(information.NA(), xlerrors.ERROR_CODE_NA)

    def test_ISNA(self):
        self.assertTrue(information.ISNA(information.NA()))
        self.assertFalse(information.ISNA(xlerrors.ValueExcelError))

    def test_ISEVEN(self):
        self.assertFalse(information.ISEVEN(func_xltypes.Number(1)))
        self.assertTrue(information.ISEVEN(func_xltypes.Number(2)))

        self.assertFalse(information.ISEVEN(-1))
        self.assertTrue(information.ISEVEN(2.5))
        self.assertFalse(information.ISEVEN(5))
        self.assertTrue(information.ISEVEN(0))
        self.assertTrue(information.ISEVEN(date.DATE(2011, 12, 23)))

    def test_ISEVEN_error(self):
        self.assertIsInstance(information.ISEVEN(func_xltypes.Text("hello")),
                              xlerrors.ValueExcelError)

    def test_ISODD(self):
        self.assertEqual(information.ISODD(1), True)
        self.assertEqual(information.ISODD(2), False)

        self.assertEqual(information.ISODD(-1), True)
        self.assertEqual(information.ISODD(2.5), False)
        self.assertEqual(information.ISODD(5), True)

    def test_ISODD_error(self):
        self.assertIsInstance(information.ISODD(func_xltypes.Text("hello")),
                              xlerrors.ValueExcelError)

    def test_ISNUMBER(self):
        self.assertFalse(information.ISNUMBER(func_xltypes.Text("hello")))
        self.assertTrue(information.ISNUMBER(func_xltypes.Number(1234)))
        self.assertFalse(information.ISNUMBER("hello"))
        self.assertTrue(information.ISNUMBER(1234))
        self.assertFalse(information.ISNUMBER("1234"))

    def test_ISERR(self):
        num_error = xlerrors.NumExcelError()
        na_error = xlerrors.NaExcelError()

        self.assertTrue(information.ISERR(num_error))
        self.assertFalse(information.ISERR(na_error))

    def test_ISERROR(self):
        num_error = xlerrors.NumExcelError()
        na_error = xlerrors.NaExcelError()

        self.assertTrue(information.ISERROR(num_error))
        self.assertTrue(information.ISERROR(na_error))
