from .. import testing


class MonthTest(testing.FunctionalTestCase):
    filename = "MONTH.xlsx"

    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual(excel_value, value)
