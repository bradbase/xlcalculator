import unittest
from xlcalculator.xlfunctions import logical, func_xltypes, xlerrors


VExpr = func_xltypes.ValueExpr


class LogicalModuleTest(unittest.TestCase):

    def test_AND(self):
        self.assertEqual(logical.AND(VExpr(3), VExpr(True), VExpr(None)), True)
        self.assertEqual(logical.AND(VExpr(True), VExpr(0)), False)

    def test_AND_without_any_args(self):
        self.assertIsInstance(logical.AND(), xlerrors.NullExcelError)

    def test_AND_with_direct_values(self):
        self.assertEqual(logical.AND(3, True, None), True)
        self.assertEqual(logical.AND(True, 0), False)

    def test_FALSE(self):
        self.assertFalse(logical.FALSE())

    def test_OR(self):
        self.assertEqual(logical.OR(VExpr(None), VExpr(3)), True)
        self.assertEqual(logical.OR(VExpr(False), VExpr(0)), False)

    def test_OR_without_any_args(self):
        self.assertIsInstance(logical.OR(), xlerrors.NullExcelError)

    def test_OR_with_dirext_values(self):
        self.assertEqual(logical.OR(3, True, None), True)
        self.assertEqual(logical.OR(False, 0), False)

    def test_IF(self):
        self.assertEqual(logical.IF(VExpr(True), VExpr(1), VExpr(2)), 1)
        self.assertEqual(logical.IF(VExpr(False), VExpr(1), VExpr(2)), 2)

    def test_IF_with_direct_values(self):
        # Values get converted to expressions to not break the APIs.
        self.assertEqual(logical.IF(True, 1, 2), 1)
        self.assertEqual(logical.IF(False, 1, 2), 2)

    def test_NOT(self):
        self.assertEqual(logical.NOT(VExpr(True)), False)
        self.assertEqual(logical.NOT(VExpr(False)), True)

    def test_NOT_with_direct_values(self):
        # Values get converted to expressions to not break the APIs.
        self.assertEqual(logical.NOT(True), False)
        self.assertEqual(logical.NOT(False), True)

    def test_TRUE(self):
        self.assertTrue(logical.TRUE())
