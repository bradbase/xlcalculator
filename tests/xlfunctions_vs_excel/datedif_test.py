from .. import testing


class DatedifTest(testing.FunctionalTestCase):
    filename = "DATEDIF.xlsx"

    def test_counta_evaluation_Y(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C2')
        value = self.evaluator.evaluate('Sheet1!C2')
        self.assertEqual(excel_value, value)

    def test_evaluation_M(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C3')
        value = self.evaluator.evaluate('Sheet1!C3')
        self.assertEqual(excel_value, value)

    def test_evaluation_D(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C4')
        value = self.evaluator.evaluate('Sheet1!C4')
        self.assertEqual(excel_value, value)

    def test_evaluation_MD(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C5')
        value = self.evaluator.evaluate('Sheet1!C5')
        self.assertEqual(excel_value, value)

    def test_evaluation_YM(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C6')
        value = self.evaluator.evaluate('Sheet1!C6')
        self.assertEqual(excel_value, value)

    def test_evaluation_YD(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C7')
        value = self.evaluator.evaluate('Sheet1!C7')
        self.assertEqual(excel_value, value)
