
import unittest
import json
from copy import deepcopy

from jsonpickle import encode, decode

from koala_xlcalculator.model.model_compiler import ModelCompiler
from koala_xlcalculator.koala_types import XLCell, XLFormula, XLRange
from koala_xlcalculator.read_excel.tokenizer import f_token
from koala_xlcalculator.model.model import Model


class TestModelCompiler(unittest.TestCase):

    def setUp(self):
        infile = open(r"./tests/resources/reader.json", "rb")
        json_bytes = infile.read()
        infile.close()
        data = decode(json_bytes, keys=True, classes=(XLCell, XLFormula, f_token, XLRange))
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        self.model = Model()
        self.model.cells = deepcopy(self.cells)
        self.model.defined_names = deepcopy(self.defined_names)
        self.model.ranges = deepcopy(self.ranges)
        self.model.formulae = deepcopy(self.formulae)


    # def teardown(self):
    #     pass


    def test_read_and_parse_archive(self):
        model_compiler = ModelCompiler()
        new_model = model_compiler.read_and_parse_archive(r"./tests/resources/reader.xlsm")

        self.assertEqual(self.model, new_model)


    def test_build_defined_names(self):
        model_compiler = ModelCompiler()
        archive = ModelCompiler.read_excel_file(r"./tests/resources/reader.xlsm")
        model_compiler.model.cells, model_compiler.model.formulae = archive.read_cells()
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()

        self.assertEqual(self.model.defined_names, model_compiler.model.defined_names)


    def test_link_cells_to_defined_names(self):
        model_compiler = ModelCompiler()
        archive = ModelCompiler.read_excel_file(r"./tests/resources/reader.xlsm")
        model_compiler.model.cells, model_compiler.model.formulae = archive.read_cells()
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()
        model_compiler.link_cells_to_defined_names()

        self.assertEqual(self.model.cells, model_compiler.model.cells)


    def test_build_ranges(self):
        model_compiler = ModelCompiler()
        archive = ModelCompiler.read_excel_file(r"./tests/resources/reader.xlsm")
        model_compiler.model.cells, model_compiler.model.formulae = archive.read_cells()
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()
        model_compiler.build_ranges()

        self.assertEqual(self.model.ranges, model_compiler.model.ranges)
