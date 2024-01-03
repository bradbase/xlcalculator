from .. import testing


class SqrtTest(testing.FunctionalTestCase):
    filename = "IFERROR.xlsx"

    def test_evaluation_C1(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!C1!'),
            self.evaluator.evaluate('Sheet1!C1')
        )

    def test_evaluation_C2(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!C2!'),
            self.evaluator.evaluate('Sheet1!C2')
        )

    def test_evaluation_C3(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!C3!'),
            self.evaluator.evaluate('Sheet1!C3')
        )

    def test_evaluation_C4(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!C4!'),
            self.evaluator.evaluate('Sheet1!C4')
        )

    def test_evaluation_C5(self):
        self.assertEqual(
            self.evaluator.get_cell_value('Sheet1!C5!'),
            self.evaluator.evaluate('Sheet1!C5')
        )
