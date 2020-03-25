
import unittest

from koala_xlcalculator.koala_types import XLRange


class TestXLRange(unittest.TestCase):

    # def setUp(self):
    #     pass

    # def teardown(self):
    #     pass


    def test_address(self):

        xlrange_00 = XLRange('Sheet1!A1:A5', 'Sheet1!A1:A5')
        address_00 = [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3'], ['Sheet1!A4'], ['Sheet1!A5']]
        self.assertEqual(address_00, xlrange_00.address)


        xlrange_01 = XLRange('Sheet1!A1:A5,Sheet1!C1:C5,Sheet1!E1:E5', 'Sheet1!A1:A5,Sheet1!C1:C5,Sheet1!E1:E5')
        address_01 = [['Sheet1!A1', 'Sheet1!C1', 'Sheet1!E1'],
                        ['Sheet1!A2', 'Sheet1!C2', 'Sheet1!E2'],
                        ['Sheet1!A3', 'Sheet1!C3', 'Sheet1!E3'],
                        ['Sheet1!A4', 'Sheet1!C4', 'Sheet1!E4'],
                        ['Sheet1!A5', 'Sheet1!C5', 'Sheet1!E5']]

        self.assertEqual(address_01, xlrange_01.cells)

        xlrange_02 = XLRange('Sheet1!A1:A5, Sheet1!C1:C5, Sheet1!E1:E5', 'Sheet1!A1:A5, Sheet1!C1:C5, Sheet1!E1:E5')
        address_02 = [['Sheet1!A1', 'Sheet1!C1', 'Sheet1!E1'],
                        ['Sheet1!A2', 'Sheet1!C2', 'Sheet1!E2'],
                        ['Sheet1!A3', 'Sheet1!C3', 'Sheet1!E3'],
                        ['Sheet1!A4', 'Sheet1!C4', 'Sheet1!E4'],
                        ['Sheet1!A5', 'Sheet1!C5', 'Sheet1!E5']
                        ]
        self.assertEqual(address_02, xlrange_02.cells)

        xlrange_03 = XLRange('Sheet1!A1:A5,C1:C5,E1:E5', 'Sheet1!A1:A5,C1:C5,E1:E5')
        address_03 = [['Sheet1!A1', 'Sheet1!C1', 'Sheet1!E1'],
                        ['Sheet1!A2', 'Sheet1!C2', 'Sheet1!E2'],
                        ['Sheet1!A3', 'Sheet1!C3', 'Sheet1!E3'],
                        ['Sheet1!A4', 'Sheet1!C4', 'Sheet1!E4'],
                        ['Sheet1!A5', 'Sheet1!C5', 'Sheet1!E5']
                        ]
        self.assertEqual(address_03, xlrange_03.address)

        xlrange_04 = XLRange('Sheet1!A1:C5', 'Sheet1!A1:C5')
        address_04 = [['Sheet1!A1', 'Sheet1!B1', 'Sheet1!C1'],
                        ['Sheet1!A2', 'Sheet1!B2', 'Sheet1!C2'],
                        ['Sheet1!A3', 'Sheet1!B3', 'Sheet1!C3'],
                        ['Sheet1!A4', 'Sheet1!B4', 'Sheet1!C4'],
                        ['Sheet1!A5', 'Sheet1!B5', 'Sheet1!C5']
                        ]
        self.assertEqual(address_04, xlrange_04.address)


    def test_cells(self):
        cells = [['Sheet1!A1'], ['Sheet1!A2'], ['Sheet1!A3'], ['Sheet1!A4'], ['Sheet1!A5']]
        xlrange = XLRange('Sheet1!A1:A5', 'Sheet1!A1:A5')
        self.assertEqual(cells, xlrange.cells)
