

from tests import testing


class CrossSheetTest(testing.FunctionalTestCase):

    filename = "cross_sheet.xlsx"

    def test_reference(self):
        """
            Tests if a reference to a cell on another sheet (Sheet1),
            which refers to a cell on it's own sheet (Sheet2) is resolved
            properly as being on Sheet2

            Also validates if after that passing in only "A1" refers back to
            Sheet1
        """
        # This tests going back and forth between original Sheet1 and Sheet2
        # multiple times in different orders
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!D2'),
            24
        )

        for coord, cell in self.model.cells.items():
            if cell.formula is None or cell.value is None:
                continue
            expected = cell.value
            cell.value = None
            got = self.evaluator.evaluate(coord)
            self.assertEqual(
                got,
                expected,
                msg=f"{coord} got: {got} expected: {expected}"
            )

        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A5'),
            4
        )
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!C5'),
            28
        )

        self.assertEqual(
            self.evaluator.evaluate('Sheet2!C1'),
            12
        )
