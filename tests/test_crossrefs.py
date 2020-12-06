
import unittest

from . import testing


class CrossSheetTest(testing.FunctionalTestCase):

    filename = "cross_sheet.xlsx"

    def test_reference(self):
        self.assertEqual(
            self.evaluator.evaluate('Sheet1!A1'),
            self.evaluator.get_cell_value('Sheet2!A1')
        )
