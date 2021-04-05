from .. import testing


class CountIfTest(testing.FunctionalTestCase):
    filename = "IF.xlsx"

    def test_evaluation_ABCDE_1(self):
        for col in "ABCDE":
            cell = f'Sheet1!{col}1'
            excel_value = self.evaluator.get_cell_value(cell)
            value = self.evaluator.evaluate(cell)
            self.assertEqual(excel_value, value)
