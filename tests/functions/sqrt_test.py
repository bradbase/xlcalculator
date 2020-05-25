import unittest
from .. import testing


class SqrtTest(testing.FunctionalTestCase):
    filename = "SQRT.xlsx"

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual( excel_value, value )

    def test_evaluation_B1(self):
        # with self.assertRaises(ExcelError):
        #     self.evaluator.evaluate('Sheet1!B1')

        with self.assertRaises(Exception) as context:
            self.evaluator.evaluate('Sheet1!B1')
            self.assertTrue('#NUM!' in context.exception)

    @unittest.skip("Can't work as we have not implemented function ABS")
    def test_evaluation_C1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C1')
        value = self.evaluator.evaluate('Sheet1!C1')
        self.assertEqual( excel_value, value )
