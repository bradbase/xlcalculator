import unittest

from xlcalculator.xlfunctions import xlerrors


class ExcelErrorTest(unittest.TestCase):

    def test_init(self):
        err = xlerrors.ExcelError('#N/A', 'Not applicable')
        self.assertEqual(err.value, '#N/A')
        self.assertEqual(err.info, 'Not applicable')

    def test_str(self):
        err = xlerrors.ExcelError('#N/A', 'Not applicable')
        self.assertEqual(str(err), '#N/A')

    def test_eq(self):
        err = xlerrors.ExcelError('#N/A', 'Not applicable')
        self.assertEqual(err, err)

    def test_eq_to_string(self):
        err = xlerrors.ExcelError('#N/A', 'Not applicable')
        self.assertEqual(err, '#N/A')


class SpecificExcelErrorTest(unittest.TestCase):

    def test_init(self):
        err = xlerrors.SpecificExcelError('Not applicable')
        self.assertEqual(err.value, None)
        self.assertEqual(err.info, 'Not applicable')


class NullExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.NullExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_NULL)


class DivZeroExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.DivZeroExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_DIV_ZERO)


class ValueExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.ValueExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_VALUE)


class RefExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.RefExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_REF)


class NameExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.NameExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_NAME)


class NumExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.NumExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_NUM)


class NaExcelErrorTest(unittest.TestCase):

    def test_error(self):
        err = xlerrors.NaExcelError('Error')
        self.assertEqual(str(err), xlerrors.ERROR_CODE_NA)
