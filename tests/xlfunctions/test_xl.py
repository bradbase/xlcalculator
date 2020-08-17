import mock
import typing
import unittest

from xlcalculator.xlfunctions import xl, xlerrors, func_xltypes


class FunctionsTest(unittest.TestCase):

    def test_register(self):
        fn = xl.Functions()

        def sample():
            pass

        fn.register(sample)
        self.assertDictEqual(dict(fn), {'sample': sample})

    def test_register_withName(self):
        fn = xl.Functions()

        def sample():
            pass

        fn.register(sample, 'mysample')
        self.assertDictEqual(dict(fn), {'mysample': sample})

    def test_getattr(self):
        fn = xl.Functions()

        def sample():
            pass

        fn.register(sample)
        self.assertEqual(fn.sample, sample)

    def test_getattr_withUnknown(self):
        fn = xl.Functions()
        with self.assertRaises(AttributeError):
            fn.sample

    def test_register_decorator(self):

        with mock.patch(
            'xlcalculator.xlfunctions.xl.FUNCTIONS', xl.Functions()
        ):

            @xl.register()
            def sample():
                pass

            self.assertDictEqual(dict(xl.FUNCTIONS), {'sample': sample})


class XlModuleTest(unittest.TestCase):

    def test_flatten(self):
        self.assertEqual(xl.flatten([1, [2, 3], [4]]), [1, 2, 3, 4])
        df = func_xltypes.Array([[1, 2], [3, 4]])
        self.assertEqual(xl.flatten(df), [1, 2, 3, 4])
        self.assertEqual(xl.flatten([df]), [1, 2, 3, 4])

    def test_length(self):
        self.assertEqual(xl.length([1, [2, 3], [4]]), 4)
        df = func_xltypes.Array([[1, 2], [3, 4]])
        self.assertEqual(xl.length(df), 4)

    def test_validate_args(self):

        @xl.validate_args
        def func(arg: func_xltypes.XlNumber):
            return arg

        self.assertEqual(func('1'), 1)

    def test_validate_args_fail(self):

        @xl.validate_args
        def func(arg: func_xltypes.XlNumber):
            return arg

        self.assertIsInstance(func('bad'), xlerrors.ValueExcelError)
        self.assertIsInstance(
            func(xlerrors.ValueExcelError()), xlerrors.ValueExcelError)

    def test_validate_args_with_list(self):

        @xl.validate_args
        def func(*args: typing.List[func_xltypes.XlNumber]):
            return args

        self.assertEqual(func('1', 2), (1, 2))

    def test_validate_args_with_union(self):

        @xl.validate_args
        def func(
            arg: typing.Union[func_xltypes.XlNumber, func_xltypes.XlText]
        ):
            return arg

        self.assertEqual(func('data'), 'data')
        self.assertEqual(func(1), 1)

    def test_validate_args_with_union_and_bad_value(self):

        @xl.validate_args
        def func(
            arg: typing.Union[func_xltypes.XlNumber, func_xltypes.XlDateTime]
        ):
            return arg

        self.assertIsInstance(func('data'), xlerrors.ValueExcelError)

    def test_validate_args_without_type(self):

        @xl.validate_args
        def func(arg):
            return arg

        obj = object()
        self.assertEqual(func(obj), obj)
