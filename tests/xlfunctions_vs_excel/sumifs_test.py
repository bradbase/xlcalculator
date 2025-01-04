from tests import testing


class SumIfsTest(testing.FunctionalTestCase):
    filename = "SUMIFS.xlsx"

    def test_evaluation_A10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual(excel_value, value)
