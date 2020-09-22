import unittest

from xlcalculator.xlfunctions import information, xlerrors, func_xltypes


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
