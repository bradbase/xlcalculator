
# Excel reference: https://support.office.com/en-us/article/npv-function-8672cb67-2576-4d07-b67b-ac28acf2a568

import unittest

import pandas as pd

from koala_xlcalculator.function_library import NPV
from koala_xlcalculator.koala_types import XLCell
from koala_xlcalculator.exceptions import ExcelError
from koala_xlcalculator import ModelCompiler
from koala_xlcalculator import Evaluator


class TestNPV(unittest.TestCase):

    def setUp(self):
        compiler = ModelCompiler()
        self.model = compiler.read_and_parse_archive(r"./tests/resources/NPV.xlsx")
        self.model.build_code()
        self.evaluator = Evaluator(self.model)

    # def teardown(self):
    #     pass


    def test_npv_basic(self):
        range_00 = pd.DataFrame([[1, 2, 3]])
        self.assertEqual(round(NPV.npv(0.06, range_00), 7), 5.2422470)
        self.assertEqual(round(NPV.npv(0.06, 1, 2, 3), 7), 5.2422470)
        self.assertEqual(round(NPV.npv(0.06, 1), 7), 0.9433962)

        self.assertEqual(round(NPV.npv(0.1, -10000, 3000, 4200, 6800), 2), 1188.44)

        range_01 = pd.DataFrame([[8000, 9200, 10000, 12000, 14500]])
        self.assertEqual(round(NPV.npv(0.08, range_01) + -40000, 2), 1922.06)
        self.assertEqual(round(NPV.npv(0.08, range_01, -9000) + -40000, 2), -3749.47)


    @unittest.skip("""Problem evalling: unsupported operand type(s) for +: 'int' and 'XLCell' for Sheet1!A1, NPV.npv(self.eval_ref("Sheet1!A2"),self.eval_ref("Sheet1!A3"),self.eval_ref("Sheet1!A4"),self.eval_ref("Sheet1!A5"),self.eval_ref("Sheet1!A6"))""")
    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )


    @unittest.skip("""Problem evalling: unsupported operand type(s) for +: 'int' and 'XLCell' for Sheet1!B1, Evaluator.apply("add",NPV.npv(self.eval_ref("Sheet1!B2"),self.eval_ref("Sheet1!B4:B8"),Evaluator.apply_one("minus", 9000, None, None)),self.eval_ref("Sheet1!B3"),None))""")
    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqual( excel_value, value )
