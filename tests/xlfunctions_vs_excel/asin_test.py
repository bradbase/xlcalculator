from tests import testing


class ASINTest(testing.FunctionalTestCase):

    filename = "ASIN.xlsx"

    def test_evaluation_A2(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A2'),
            self.evaluator.get_cell_value('Sheet1!A2')
        )

    def test_evaluation_A3(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A3'),
            self.evaluator.get_cell_value('Sheet1!A3')
        )

    def test_evaluation_A4(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A4'),
            self.evaluator.get_cell_value('Sheet1!A4')
        )

    def test_evaluation_A5(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A5'),
            self.evaluator.get_cell_value('Sheet1!A5')
        )
