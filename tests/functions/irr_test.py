import unittest
from .. import testing


class IRRTest(testing.FunctionalTestCase):
    filename = "IRR.xlsx"

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqualTruncated( excel_value, value )

    def test_evaluation_B1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B1')
        value = self.evaluator.evaluate('Sheet1!B1')
        self.assertEqualTruncated( excel_value, value, 13 )

    @unittest.skip("""Problem evalling: guess value for excellib.irr() is #N/A and not 0 for Sheet1!C1, IRR.irr(self.eval_ref("Sheet1!A2:A4"),Evaluator.apply_one("minus", 0.1, None, None))""")
    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual( excel_value, value )
