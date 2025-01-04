from tests import testing


class ChooseTest(testing.FunctionalTestCase):
    filename = "choose.xlsx"

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value_00 = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value, value_00)

    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value_01 = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual(excel_value, value_01)

    def test_evaluation_A3(self):
        excel_value = [[1, 2, 3]]
        value_00 = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual(excel_value, value_00.values.tolist())
