from tests import testing

from xlcalculator.xlfunctions import xlerrors


class SqrtTest(testing.FunctionalTestCase):
    filename = "SQRT.xlsx"

    def test_evaluation_A1(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!A1'),
            self.evaluator.evaluate('Sheet1!A1')
        )

    def test_evaluation_B1(self):
        self.assertIsInstance(
            self.evaluator.evaluate('Sheet1!B1'), xlerrors.NumExcelError)

    def test_evaluation_C1(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!C1'),
            self.evaluator.evaluate('Sheet1!C1')
        )
