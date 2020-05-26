from .. import testing


class LogicalFunctionsTest(testing.FunctionalTestCase):

    filename = "logical.xlsx"

    def test_IF_false_case(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!C2'),
            self.evaluator.get_cell_value('Sheet1!C2')
        )

    def test_IF_true_case(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!C3'),
            self.evaluator.get_cell_value('Sheet1!C3')
        )

    def test_AND_false_case(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!D2'),
            self.evaluator.get_cell_value('Sheet1!D2')
        )

    def test_AND_true_case(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!D3'),
            self.evaluator.get_cell_value('Sheet1!D3')
        )

    def test_OR_false_case(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!E2'),
            self.evaluator.get_cell_value('Sheet1!E2')
        )

    def test_OR_true_case(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!E3'),
            self.evaluator.get_cell_value('Sheet1!E3')
        )
