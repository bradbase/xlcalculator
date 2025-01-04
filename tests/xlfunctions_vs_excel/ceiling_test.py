from tests import testing


class CEILINGTest(testing.FunctionalTestCase):

    filename = "CEILING.xlsx"

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

    def test_evaluation_A6(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A6'),
            self.evaluator.get_cell_value('Sheet1!A6')
        )

    def test_evaluation_A7(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A7'),
            self.evaluator.get_cell_value('Sheet1!A7')
        )

    def test_evaluation_A8(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A8'),
            self.evaluator.get_cell_value('Sheet1!A8')
        )

    def test_evaluation_A9(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A9'),
            self.evaluator.get_cell_value('Sheet1!A9')
        )
