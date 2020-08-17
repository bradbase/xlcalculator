from .. import testing


class VLookupTest(testing.FunctionalTestCase):
    filename = "VLOOKUP.xlsx"

    def test_evaluation_B7(self):
        # Range Match - Unsupposrted
        with self.assertRaises(RuntimeError):
            self.evaluator.evaluate('Sheet1!B7')

    def test_evaluation_E7(self):
        # Exact Match.
        excel_value = self.evaluator.get_cell_value('Sheet1!E7')
        value = self.evaluator.evaluate('Sheet1!E7')
        self.assertEqual(excel_value, value)
