import jsonpickle
import unittest

from xlcalculator import reader, xltypes, tokenizer
from . import testing


class ReaderTest(unittest.TestCase):

    def setUp(self):
        infile = open(testing.get_resource("reader.json"), "rb")
        json_bytes = infile.read()
        infile.close()
        data = jsonpickle.decode(
            json_bytes, keys=True,
            classes=(
                xltypes.XLCell, xltypes.XLFormula, xltypes.XLRange,
                tokenizer.f_token
            )
        )
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

    def test_read_cells(self):
        archive = reader.Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        cells, formulae, ranges = \
            archive.read_cells(ignore_sheets=['Eleventh'])

        self.assertEqual(sorted(self.cells.keys()), sorted(cells.keys()))

    @unittest.skip(
        "Need to re-create reader.json as some elements of dynamic arrays "
        "are not being read correctly")
    def test_read_formulae(self):
        archive = reader.Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        cells, formulae, ranges = \
            archive.read_cells(ignore_sheets=['Eleventh'])

        self.assertEqual(sorted(self.formulae.keys()), sorted(formulae.keys()))

    def test_read_defined_names(self):
        archive = reader.Reader(testing.get_resource("reader.xlsm"))
        archive.read()
        defined_names = archive.read_defined_names()

        its_a_blank = (
            set(sorted(defined_names.keys()))
            - set(sorted(self.defined_names.keys()))
        )
        self.assertEqual(list(its_a_blank), ['Its_a_blank'])
