from .. import testing


class SumTest(testing.FunctionalTestCase):
    filename = "BASES.xlsx"

    def test_evalMatrix(self):
        for col in 'BCDEFGHIJKLM':
            for row in range(4, 20):
                addr = f'Sheet1!{col}{row}'
                excel_value = self.evaluator.get_cell_value(addr)
                value = self.evaluator.evaluate(addr)
                self.assertEqual(excel_value, value, addr)

    def test_converstionWithReferences(self):
        excel_value = self.evaluator.get_cell_value('Sheet1!R14')
        value = self.evaluator.evaluate('Sheet1!R14')
        self.assertEqual(excel_value, value)
