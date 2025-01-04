from tests import testing


class YearTest(testing.FunctionalTestCase):
    filename = "YEAR.xlsx"

    def test_evaluation_A5(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A5')
        value = self.evaluator.evaluate('Sheet1!A5')
        self.assertEqual(excel_value, value)

    def test_evaluation_A6(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A6')
        value = self.evaluator.evaluate('Sheet1!A6')
        self.assertEqual(excel_value, value)

    def test_evaluation_C5(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C5')
        value = self.evaluator.evaluate('Sheet1!C5')
        self.assertEqual(excel_value, value)

    def test_evaluation_D5(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D5')
        value = self.evaluator.evaluate('Sheet1!D5')
        self.assertEqual(excel_value, value)
