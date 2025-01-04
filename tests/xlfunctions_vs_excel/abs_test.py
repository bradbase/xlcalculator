from tests import testing


class ABSTest(testing.FunctionalTestCase):

    filename = "ABS.xlsx"

    def test_evaluation_A1(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A1'),
            self.evaluator.get_cell_value('Sheet1!A1')
        )
