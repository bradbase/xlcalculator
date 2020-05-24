
import unittest
import json

from jsonpickle import encode, decode

from xlcalculator.read_excel.reader import Reader
from xlcalculator.types import XLCell, XLFormula, XLRange
from xlcalculator.read_excel.tokenizer import f_token

from ..function_library import testing


class TestRead_excel(unittest.TestCase):

    def setUp(self):
        infile = open(testing.get_resource("reader.json"), "rb")
        json_bytes = infile.read()
        infile.close()
        data = decode(
            json_bytes, keys=True,
            classes=(XLCell, XLFormula, f_token, XLRange))
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

    def test_read_cells(self):
        archive = Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        cells, formulae, ranges = archive.read_cells(ignore_sheets=['Eleventh'])

        self.assertEqual(sorted(self.cells.keys()), sorted(cells.keys()))

    @unittest.skip(
        "Need to re-create reader.json as some elements of dynamic arrays "
        "are not being read correctly")
    def test_read_formulae(self):
        archive = Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        cells, formulae, ranges = archive.read_cells(ignore_sheets=['Eleventh'])

        self.assertEqual(sorted(self.formulae.keys()), sorted(formulae.keys()))

    def test_read_defined_names(self):
        archive = Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        defined_names = archive.read_defined_names()

        its_a_blank = (
            set(sorted(defined_names.keys())) -
            set(sorted(self.defined_names.keys()))
        )
        self.assertEqual(list(its_a_blank), ['Its_a_blank'])
