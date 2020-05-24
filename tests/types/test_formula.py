import unittest

from xlcalculator.types import XLFormula
from xlcalculator.read_excel.tokenizer import f_token

from ..xlcalculator_test import XlCalculatorTestCase


class XLFormulaTest(XlCalculatorTestCase):

    def test_formula(self):
        self.assertEqual(XLFormula('=SUM(A1:B1)').formula, '=SUM(A1:B1)')

    def test_tokens(self):
        self.assertASTNodesEqual(
            XLFormula('=SUM(A1:B1)').tokens,
            [
                f_token(tvalue='SUM', ttype='function', tsubtype='start'),
                f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'),
                f_token(tvalue='', ttype='function', tsubtype='stop')
            ]
        )
