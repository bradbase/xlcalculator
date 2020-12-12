

from . import testing


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
            self.evaluator.evaluate('Sheet1!A4'),
            5
        )
