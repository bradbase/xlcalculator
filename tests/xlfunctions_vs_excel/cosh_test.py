from tests import testing


class COSHTest(testing.FunctionalTestCase):
    filename = "COSH.xlsx"

    def test_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertAlmostEqual(excel_value, value, 8)

    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertAlmostEqual(excel_value, value)
