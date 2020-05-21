
import unittest
import json
from copy import deepcopy

from jsonpickle import encode, decode

from xlcalculator.model.model_compiler import ModelCompiler
from xlcalculator.xlcalculator_types import XLCell, XLFormula, XLRange
from xlcalculator.read_excel.tokenizer import f_token
from xlcalculator.model.model import Model


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
        new_model = model_compiler.read_and_parse_archive(r"./tests/resources/reader.xlsm", ignore_sheets=['Eleventh'])

        self.assertEqual(self.model, new_model)


    def test_build_defined_names(self):
        model_compiler = ModelCompiler()
        archive = model_compiler.read_excel_file(r"./tests/resources/reader.xlsm")
        model_compiler.model.cells, model_compiler.model.formulae, model_compiler.model.ranges = archive.read_cells(ignore_sheets=['Eleventh'])
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()

        self.assertEqual(self.model.defined_names, model_compiler.model.defined_names)


    def test_link_cells_to_defined_names(self):
        model_compiler = ModelCompiler()
        archive = model_compiler.read_excel_file(r"./tests/resources/reader.xlsm")
        model_compiler.model.cells, model_compiler.model.formulae, model_compiler.model.ranges = archive.read_cells(ignore_sheets=['Eleventh'])
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()
        model_compiler.link_cells_to_defined_names()

        self.assertEqual(self.model.cells, model_compiler.model.cells)


    def test_build_ranges(self):
        model_compiler = ModelCompiler()
        archive = model_compiler.read_excel_file(r"./tests/resources/reader.xlsm")
        model_compiler.model.cells, model_compiler.model.formulae, model_compiler.model.ranges = archive.read_cells(ignore_sheets=['Eleventh'])
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()
        model_compiler.build_ranges()

        self.assertEqual(self.model.ranges, model_compiler.model.ranges)


    def test_extract_cells(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(r"./tests/resources/reader.xlsm", ignore_sheets=['Eleventh'])
        extracted_model = ModelCompiler.extract(reader_model, focus=['First!A2', 'First!B2', 'First!C2'])

        reference_model = Model()
        reference_model.set_cell_value('First!A2', 0.1)
        reference_model.set_cell_value('First!B2', 0.2)
        reference_model.set_cell_value('First!C2', 0.3)

        self.assertEqual(reference_model.cells, extracted_model.cells)


    def test_extract_defined_names(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(r"./tests/resources/reader.xlsm", ignore_sheets=['Eleventh'])
        extracted_model = ModelCompiler.extract(reader_model, focus=['Hundred', 'My_Range'])

        reference_model = Model()
        reference_model.set_cell_value('Eighth!B1', 100)
        reference_model.set_cell_value('Eighth!A1', 1)
        reference_model.set_cell_value('Eighth!A2', 2)
        reference_model.set_cell_value('Eighth!A3', 3)
        reference_model.set_cell_value('Eighth!A4', 4)
        reference_model.set_cell_value('Eighth!A5', 5)
        reference_model.set_cell_value('Eighth!A6', 6)
        reference_model.set_cell_value('Eighth!A7', 7)
        reference_model.set_cell_value('Eighth!A8', 8)
        reference_model.set_cell_value('Eighth!A9', 9)
        reference_model.set_cell_value('Eighth!A10', 10)
        reference_model.defined_names['Hundred'] = deepcopy(reader_model.defined_names['Hundred'])
        reference_model.defined_names['My_Range'] = deepcopy(reader_model.defined_names['My_Range'])

        self.assertEqual(reference_model.cells, extracted_model.cells)
        self.assertEqual(reference_model.defined_names, extracted_model.defined_names)
        self.assertEqual(reference_model, extracted_model)


    def test_extract(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(r"./tests/resources/reader.xlsm", ignore_sheets=['Eleventh'])
        extracted_model = ModelCompiler.extract(reader_model,
                                                focus=['First!A2',
                                                    'First!B2',
                                                    'First!C2',
                                                    'Fourth!A2',
                                                    'Hundred',
                                                    'My_Range'],
                                                )

        reference_model = Model()
        reference_model.set_cell_value('First!A2', 0.1)
        reference_model.set_cell_value('First!B2', 0.2)
        reference_model.set_cell_value('First!C2', 0.3)
        reference_model.set_cell_value('Eighth!B1', 100)
        reference_model.set_cell_value('Eighth!A1', 1)
        reference_model.set_cell_value('Eighth!A2', 2)
        reference_model.set_cell_value('Eighth!A3', 3)
        reference_model.set_cell_value('Eighth!A4', 4)
        reference_model.set_cell_value('Eighth!A5', 5)
        reference_model.set_cell_value('Eighth!A6', 6)
        reference_model.set_cell_value('Eighth!A7', 7)
        reference_model.set_cell_value('Eighth!A8', 8)
        reference_model.set_cell_value('Eighth!A9', 9)
        reference_model.set_cell_value('Eighth!A10', 10)
        reference_model.defined_names['Hundred'] = deepcopy(reader_model.defined_names['Hundred'])
        reference_model.defined_names['My_Range'] = deepcopy(reader_model.defined_names['My_Range'])
        reference_model.cells['Fourth!A2'] = reader_model.cells['Fourth!A2']

        self.assertEqual(reference_model.cells, extracted_model.cells)
        self.assertEqual(reference_model.defined_names, extracted_model.defined_names)
