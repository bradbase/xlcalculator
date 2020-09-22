from .. import testing


class MatchTest(testing.FunctionalTestCase):
    filename = "MATCH.xlsx"

    def test_evaluation_A8(self):
        # Match defaults.
        excel_value = self.evaluator.get_cell_value('Sheet1!A8')
        value = self.evaluator.evaluate('Sheet1!A8')
        self.assertEqual(excel_value, value)

    def test_evaluation_A9(self):
        # Match defaults.
        excel_value = self.evaluator.get_cell_value('Sheet1!A9')
        value = self.evaluator.evaluate('Sheet1!A9')
        self.assertEqual(excel_value, value)

    def test_evaluation_A10(self):
        # Match defaults.
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual(excel_value, value)
