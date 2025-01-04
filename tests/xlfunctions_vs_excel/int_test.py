from tests import testing


class IntTest(testing.FunctionalTestCase):
    filename = "INT.xlsx"

    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual(excel_value, value)
