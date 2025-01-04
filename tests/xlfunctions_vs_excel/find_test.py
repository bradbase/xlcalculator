from tests import testing


class FindTest(testing.FunctionalTestCase):
    filename = "FIND.xlsx"

    def test_evaluation_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value, value)

    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual(excel_value, value)

    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual(excel_value, value)

    def test_evaluation_MID_FIND_A1(self):
        excel_value = self.evaluator.get_cell_value('Sheet2!A1')
        value = self.evaluator.evaluate('Sheet2!A1')
        self.assertEqual(excel_value, value)

    def test_evaluation_MID_FIND_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet2!A2')
        value = self.evaluator.evaluate('Sheet2!A2')
        self.assertEqual(excel_value, value)

    def test_evaluation_MID_FIND_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet2!A3')
        value = self.evaluator.evaluate('Sheet2!A3')
        self.assertEqual(excel_value, value)
