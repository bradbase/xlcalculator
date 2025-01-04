from tests import testing


class DayTest(testing.FunctionalTestCase):
    filename = "DAY.xlsx"

    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual(excel_value, value)
