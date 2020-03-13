
import unittest
import json

from jsonpickle import encode, decode

from koala_xlcalculator.read_excel.reader import Reader
from koala_xlcalculator.koala_types import XLCell, XLFormula, XLRange
from koala_xlcalculator.read_excel.tokenizer import f_token


class TestRead_excel(unittest.TestCase):

    def setUp(self):
        infile = open(r"./tests/read_excel/reader.json", "rb")
        json_bytes = infile.read()
        infile.close()
        data = decode(json_bytes, keys=True, classes=(XLCell, XLFormula, f_token, XLRange))
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        self.shared_strings_metadata = {0: 'Header 1', 1: 'Header 2', 2: 'Header 3', 3: 'Header 4', 4: 'Header 5', 5: 'Head 1', 6: 'Head 2', 7: 'Head 3'}

        self.defined_name_metadata = {'Hundred': 'Eighth!$B$1', 'My_Range': 'Eighth!$A$1:$A$10', 'Split_range': 'Eighth!$A$1:$A$10,Eighth!$C$1:$C$10'}

        self.worksheet_metadata = {'rId1': {'name': 'First', 'sheetId': '1', 'Target': 'worksheets/sheet1.xml'}, 'rId2': {'name': 'Second', 'sheetId': '2', 'Target': 'worksheets/sheet2.xml'}, 'rId3': {'name': 'Third', 'sheetId': '3', 'Target': 'worksheets/sheet3.xml'}, 'rId4': {'name': 'Fourth', 'sheetId': '4', 'Target': 'worksheets/sheet4.xml'}, 'rId5': {'name': 'Fifth', 'sheetId': '5', 'Target': 'worksheets/sheet5.xml'}, 'rId6': {'name': 'Sixth', 'sheetId': '6', 'Target': 'worksheets/sheet6.xml'}, 'rId7': {'name': 'Seventh', 'sheetId': '7', 'Target': 'worksheets/sheet7.xml'}, 'rId8': {'name': 'Eighth', 'sheetId': '8', 'Target': 'worksheets/sheet8.xml'}, 'rId9': {'name': 'Ninth', 'sheetId': '9', 'Target': 'worksheets/sheet9.xml'}, 'rId10': {'name': 'Tenth', 'sheetId': '10', 'Target': 'worksheets/sheet10.xml'}}


    # def teardown(self):
    #     pass


    def test_read_cells(self):
        archive = Reader(r"./tests/read_excel/reader.xlsm")
        archive.read()
        cells, formulae = archive.read_cells()

        self.assertEqual(sorted(self.cells.keys()), sorted(cells.keys()))


    def test_read_formulae(self):
        archive = Reader(r"./tests/read_excel/reader.xlsm")
        archive.read()
        cells, formulae = archive.read_cells()

        self.assertEqual(sorted(self.formulae.keys()), sorted(formulae.keys()))


    def test_read_defined_names(self):
        archive = Reader(r"./tests/read_excel/reader.xlsm")
        archive.read()
        defined_names = archive.read_defined_names()

        self.assertEqual(sorted(self.defined_names.keys()), sorted(defined_names.keys()))


    def test_build_shared_string_metadata(self):
        archive = Reader(r"./tests/read_excel/reader.xlsm")
        archive.read()
        archive.build_shared_string_metadata()

        self.assertEqual(self.shared_strings_metadata, archive.shared_strings_metadata)


    def test_build_defined_name_metadata(self):
        archive = Reader(r"./tests/read_excel/reader.xlsm")
        archive.read()
        archive.build_defined_name_metadata()

        self.assertEqual(self.defined_name_metadata, archive.defined_name_metadata)


    def test_build_worksheet_metadata(self):
        archive = Reader(r"./tests/read_excel/reader.xlsm")
        archive.read()
        archive.build_worksheet_metadata()

        self.assertEqual(self.worksheet_metadata, archive.worksheet_metadata)
