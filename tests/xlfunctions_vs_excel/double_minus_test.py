from .. import testing


class DoubleMinusTest(testing.FunctionalTestCase):
    filename = "double_minus.xlsx"

    def test_evaluation_double_unary_A_eq_B(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A1')
        value = self.evaluator.evaluate('Sheet1!A1')
        self.assertEqual(excel_value, value)

    def test_evaluation_double_unary_number(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual(excel_value, value)

    def test_evaluation_double_unary_formula(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual(excel_value, value)
