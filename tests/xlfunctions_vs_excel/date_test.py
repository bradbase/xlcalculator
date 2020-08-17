from .. import testing


class DateTest(testing.FunctionalTestCase):
    filename = "DATE.xlsx"

    def test_counta_evaluation_A2(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A2')
        value = self.evaluator.evaluate('Sheet1!A2')
        self.assertEqual(excel_value, value)

    def test_evaluation_A3(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A3')
        value = self.evaluator.evaluate('Sheet1!A3')
        self.assertEqual(excel_value, value)

    def test_evaluation_A4(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A4')
        value = self.evaluator.evaluate('Sheet1!A4')
        self.assertEqual(excel_value, value)

    def test_evaluation_A5(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A5')
        value = self.evaluator.evaluate('Sheet1!A5')
        self.assertEqual(excel_value, value)

    def test_evaluation_A6(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A6')
        value = self.evaluator.evaluate('Sheet1!A6')
        self.assertEqual(excel_value, value)

    def test_evaluation_A7(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A7')
        value = self.evaluator.evaluate('Sheet1!A7')
        self.assertEqual(excel_value, value)

    def test_evaluation_A9(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A9')
        value = self.evaluator.evaluate('Sheet1!A9')
        self.assertEqual(excel_value, value)

    def test_evaluation_A10(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A10')
        value = self.evaluator.evaluate('Sheet1!A10')
        self.assertEqual(excel_value, value)

    def test_evaluation_A11(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A11')
        value = self.evaluator.evaluate('Sheet1!A11')
        self.assertEqual(excel_value, value)

    def test_evaluation_A12(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A12')
        value = self.evaluator.evaluate('Sheet1!A12')
        self.assertEqual(excel_value, value)

    def test_evaluation_A13(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A13')
        value = self.evaluator.evaluate('Sheet1!A13')
        self.assertEqual(excel_value, value)

    def test_evaluation_A14(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A14')
        value = self.evaluator.evaluate('Sheet1!A14')
        self.assertEqual(excel_value, value)

    def test_evaluation_A15(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A15')
        value = self.evaluator.evaluate('Sheet1!A15')
        self.assertEqual(excel_value, value)

    def test_evaluation_A16(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A16')
        value = self.evaluator.evaluate('Sheet1!A16')
        self.assertEqual(excel_value, value)

    def test_evaluation_A17(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A17')
        value = self.evaluator.evaluate('Sheet1!A17')
        self.assertEqual(excel_value, value)

    def test_evaluation_A18(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A18')
        value = self.evaluator.evaluate('Sheet1!A18')
        self.assertEqual(excel_value, value)

    def test_evaluation_A19(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A19')
        value = self.evaluator.evaluate('Sheet1!A19')
        self.assertEqual(excel_value, value)

    def test_evaluation_A25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!A25')
        value = self.evaluator.evaluate('Sheet1!A25')
        self.assertEqual(excel_value, value)

    def test_evaluation_B25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!B25')
        value = self.evaluator.evaluate('Sheet1!B25')
        self.assertEqual(excel_value, value)

    def test_evaluation_C25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!C25')
        value = self.evaluator.evaluate('Sheet1!C25')
        self.assertEqual(excel_value, value)

    def test_evaluation_D25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!D25')
        value = self.evaluator.evaluate('Sheet1!D25')
        self.assertEqual(excel_value, value)

    def test_evaluation_E25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!E25')
        value = self.evaluator.evaluate('Sheet1!E25')
        self.assertEqual(excel_value, value)

    def test_evaluation_F25(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!F25')
        value = self.evaluator.evaluate('Sheet1!F25')
        self.assertEqual(excel_value, value)
