from tests import testing


class SLNTest(testing.FunctionalTestCase):
    filename = "SLN.xlsx"

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value, value)
