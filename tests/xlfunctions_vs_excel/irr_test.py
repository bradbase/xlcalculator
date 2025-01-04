from tests import testing


class IRRTest(testing.FunctionalTestCase):
    filename = "IRR.xlsx"

    def test_evaluation_A1(self):
        self.assertAlmostEqual(
            self.evaluator.get_cell_value('Sheet1!A1'),
            self.evaluator.evaluate('Sheet1!A1')
        )

    def test_evaluation_B1(self):
        self.assertAlmostEqual(
            self.evaluator.get_cell_value('Sheet1!B1'),
            self.evaluator.evaluate('Sheet1!B1')
        )

    def test_evaluation_C1(self):
        self.assertAlmostEqual(
            self.evaluator.get_cell_value('Sheet1!C1'),
            self.evaluator.evaluate('Sheet1!C1')
        )
