from tests import testing


class InformationTest(testing.FunctionalTestCase):
    filename = "INFORMATION.xlsx"

    def test_isblank_evaluation_A10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual(excel_value, value)

    def test_isblank_evaluation_B10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B10')
        value = self.evaluator.evaluate('Sheet1!B10')
        self.assertEqual(excel_value, value)

    def test_isblank_evaluation_C10(self):
        # C10 has a space.
        excel_value = self.evaluator.get_cell_value('Sheet1!C10')
        value = self.evaluator.evaluate('Sheet1!C10')
        self.assertEqual(excel_value, value)

    def test_istext_evaluation_D10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D10')
        value = self.evaluator.evaluate('Sheet1!D10')
        self.assertEqual(excel_value, value)

    def test_istext_evaluation_E10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!E10')
        value = self.evaluator.evaluate('Sheet1!E10')
        self.assertEqual(excel_value, value)

    def test_istext_evaluation_F10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!F10')
        value = self.evaluator.evaluate('Sheet1!F10')
        self.assertEqual(excel_value, value)

    def test_na_evaluation_G1(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!G1')
        value = self.evaluator.evaluate('Sheet1!G1')
        self.assertEqual(excel_value, value)

    def test_isna_evaluation_G10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!G10')
        value = self.evaluator.evaluate('Sheet1!G10')
        self.assertEqual(excel_value, value)

    def test_isna_evaluation_H10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!H10')
        value = self.evaluator.evaluate('Sheet1!H10')
        self.assertEqual(excel_value, value)
