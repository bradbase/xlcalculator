from .. import testing


class SumProductTest(testing.FunctionalTestCase):
    filename = "SUMPRODUCT.xlsx"

    def test_evaluation_D7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D7')
        value = self.evaluator.evaluate('Sheet1!D7')
        self.assertEqual(excel_value, value)
