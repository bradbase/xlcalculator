import unittest

from xlcalculator.xlcalculator_types import XLFormula
from xlcalculator.read_excel.tokenizer import f_token

from ..xlcalculator_test import XlCalculatorTestCase


class TestFormula(XlCalculatorTestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass

    def test_formula(self):

        xlformula_00 = XLFormula('=SUM(A1:B1)')
        formula_00 = '=SUM(A1:B1)'
        self.assertEqual(formula_00, xlformula_00.formula)


    def test_formula(self):

        xlformula_00 = XLFormula('=SUM(A1:B1)')
        tokens_00 = [
            f_token(tvalue='SUM', ttype='function', tsubtype='start'),
            f_token(tvalue='A1:B1', ttype='operand', tsubtype='range'),
            f_token(tvalue='', ttype='function', tsubtype='stop')
        ]
        self.assertASTNodesEqual(tokens_00, xlformula_00.tokens)


    @unittest.skip("Python code is currently built in Model.build_code() so needs to be tested in TestModel")
    def test_python_code(self):
        # TODO: see if it makes more sense to build the Python code in XLFormula.__post_init__()
        pass
