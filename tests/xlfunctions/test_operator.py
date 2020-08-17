import unittest

from xlcalculator.xlfunctions import xlerrors, operator


class OperatorModuleTest(unittest.TestCase):

    def test_OP_MUL(self):
        self.assertEqual(operator.OP_MUL(1, 2), 2)

    def test_OP_DIV(self):
        self.assertEqual(operator.OP_DIV(2, 1), 2)

    def test_OP_DIV_by_zero(self):
        self.assertIsInstance(
            operator.OP_DIV(2, 0), xlerrors.DivZeroExcelError)

    def test_OP_ADD(self):
        self.assertEqual(operator.OP_ADD(1, 2), 3)

    def test_OP_SUB(self):
        self.assertEqual(operator.OP_SUB(1, 2), -1)

    def test_OP_EQ(self):
        self.assertEqual(operator.OP_EQ(2, 2), True)
        self.assertEqual(operator.OP_EQ("2", 2), False)
        self.assertEqual(operator.OP_EQ(True, 1), True)
        self.assertEqual(operator.OP_EQ(1, 2), False)

    def test_OP_NE(self):
        self.assertEqual(operator.OP_NE(2, 2), False)
        self.assertEqual(operator.OP_NE("2", 2), True)
        self.assertEqual(operator.OP_NE(True, 1), False)
        self.assertEqual(operator.OP_NE(1, 2), True)

    def test_OP_GT(self):
        self.assertEqual(operator.OP_GT(2, 1), True)
        self.assertEqual(operator.OP_GT("2", 1), True)
        self.assertEqual(operator.OP_GT(1, 2), False)

    def test_OP_GT_with_None(self):
        self.assertEqual(operator.OP_GT(None, 2), False)
        self.assertEqual(operator.OP_GT(2, None), False)

    def test_OP_LT(self):
        self.assertEqual(operator.OP_LT(2, 1), False)
        self.assertEqual(operator.OP_LT("2", 1), False)
        self.assertEqual(operator.OP_LT(1, 2), True)

    def test_OP_LT_with_None(self):
        self.assertEqual(operator.OP_LT(None, 2), False)
        self.assertEqual(operator.OP_LT(2, None), False)

    def test_OP_GE(self):
        self.assertEqual(operator.OP_GE(2, 1), True)
        self.assertEqual(operator.OP_GE(1, 1), True)
        self.assertEqual(operator.OP_GE(1, 2), False)

    def test_OP_GE_with_None(self):
        self.assertEqual(operator.OP_GE(None, 2), False)
        self.assertEqual(operator.OP_GE(2, None), False)

    def test_OP_LE(self):
        self.assertEqual(operator.OP_LE(2, 1), False)
        self.assertEqual(operator.OP_LE(1, 1), True)
        self.assertEqual(operator.OP_LE(1, 2), True)

    def test_OP_LE_with_None(self):
        self.assertEqual(operator.OP_LE(None, 2), False)
        self.assertEqual(operator.OP_LE(2, None), False)

    def test_OP_NEG(self):
        self.assertEqual(operator.OP_NEG(1), -1)
        self.assertEqual(operator.OP_NEG(-1), 1)

    def test_OP_PERCENT(self):
        self.assertEqual(operator.OP_PERCENT(2), 0.02)
