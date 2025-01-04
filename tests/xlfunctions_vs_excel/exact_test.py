from tests import testing


class ExactTest(testing.FunctionalTestCase):
    filename = "EXACT.xlsx"

    def test_evaluation_C2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C2')
        value = self.evaluator.evaluate('Sheet1!C2')
        self.assertEqual(excel_value, value)

    def test_evaluation_C3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C3')
        value = self.evaluator.evaluate('Sheet1!C3')
        self.assertEqual(excel_value, value)

    def test_evaluation_C4(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C4')
        value = self.evaluator.evaluate('Sheet1!C4')
        self.assertEqual(excel_value, value)

    def test_evaluation_C6(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C6')
        value = self.evaluator.evaluate('Sheet1!C6')
        self.assertEqual(excel_value, value)
