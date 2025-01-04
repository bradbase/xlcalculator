from tests import testing


class ConcatenateTest(testing.FunctionalTestCase):
    filename = "CONCATENATE.xlsx"

    def test_evaluation_A6(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A6')
        value = self.evaluator.evaluate('Sheet1!A6')
        self.assertEqual(excel_value, value)
