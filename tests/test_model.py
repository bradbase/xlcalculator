
import unittest
from copy import deepcopy

from jsonpickle import decode

from xlcalculator.model import Model, ModelCompiler
from xlcalculator.xltypes import XLCell, XLFormula, XLRange
from xlcalculator.tokenizer import f_token
from xlcalculator import Evaluator

from tests import testing


class ModelTest(testing.XlCalculatorTestCase):

    def setUp(self):
        infile = open(testing.get_resource("model.json"), "rb")
        json_bytes = infile.read()
        infile.close()
        data = decode(
            json_bytes, keys=True,
            classes=(XLCell, XLFormula, f_token, XLRange))
        self.cells = data['cells']
        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        self.model = Model()
        self.model.cells = deepcopy(self.cells)
        self.model.defined_names = deepcopy(self.defined_names)
        self.model.ranges = deepcopy(self.ranges)
        self.model.formulae = deepcopy(self.formulae)

    def test_build_code(self):

        this_model = Model()
        this_model.cells = self.cells
        this_model.defined_names = self.defined_names
        this_model.ranges = self.ranges
        this_model.formulae = self.formulae

        this_model.build_code()

        self.assertEqual(self.model, this_model)

    def test_set_value(self):
        this_model = deepcopy(self.model)

        # Cells we know are in the model.

        this_model.set_cell_value('First!A2', 88)
        cell_value_00 = 88
        self.assertEqual(cell_value_00, this_model.cells['First!A2'].value)

        this_model.set_cell_value(XLCell('First!A2', 22), 0.1)
        cell_value_01 = 0.1
        self.assertEqual(cell_value_01, this_model.cells['First!A2'].value)

        this_model.set_cell_value('First!A1', 'Header_1')
        cell_value_00 = 'Header_1'
        self.assertEqual(cell_value_00, this_model.cells['First!A1'].value)

        # Cells we know are not in the model, they get added

        this_model.set_cell_value('Sheet1!A1', 88)
        cell_value_00 = 88
        self.assertEqual(cell_value_00, this_model.cells['Sheet1!A1'].value)

        this_model.set_cell_value(XLCell('Sheet1!A1', 22), 0.1)
        cell_value_01 = 0.1
        self.assertEqual(cell_value_01, this_model.cells['Sheet1!A1'].value)

    def test_get_value(self):
        this_model = deepcopy(self.model)

        # cells we know are in the model
        get_cell_value_00 = this_model.get_cell_value('First!A2')
        cell_value_00 = 0.1
        self.assertEqual(cell_value_00, get_cell_value_00)

        get_cell_value_01 = this_model.get_cell_value('First!A1')
        cell_value_01 = 'Header 1'
        self.assertEqual(cell_value_01, get_cell_value_01)

        get_cell_value_02 = this_model.get_cell_value(XLCell('First!A1', 22))
        cell_value_02 = 'Header 1'
        self.assertEqual(cell_value_02, get_cell_value_02)

        get_cell_value_03 = this_model.get_cell_value(XLCell('First!A2', 22))
        cell_value_03 = 0.1
        self.assertEqual(cell_value_03, get_cell_value_03)

        # Cells we know are not in the model, all default to the value 0

        get_cell_value_04 = this_model.get_cell_value('Sheet1!A1')
        cell_value_04 = 0
        self.assertEqual(cell_value_04, get_cell_value_04)


class ModelCompilerTest(unittest.TestCase):
    maxDiff = None

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

        self.model = Model()
        self.model.cells = deepcopy(self.cells)
        self.model.defined_names = deepcopy(self.defined_names)
        self.model.ranges = deepcopy(self.ranges)
        self.model.formulae = deepcopy(self.formulae)

    def test_read_and_parse_archive(self):
        model_compiler = ModelCompiler()
        new_model = model_compiler.read_and_parse_archive(
            testing.get_resource("reader.xlsm"), ignore_sheets=['Eleventh'])
        self.assertEqual(self.model, new_model)

    def test_build_defined_names(self):
        model_compiler = ModelCompiler()
        archive = model_compiler.read_excel_file(
            testing.get_resource("reader.xlsm"))
        (
            model_compiler.model.cells,
            model_compiler.model.formulae,
            model_compiler.model.ranges
        ) = archive.read_cells(ignore_sheets=['Eleventh'])
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()

        self.assertEqual(
            self.model.defined_names,
            model_compiler.model.defined_names
        )

    def test_link_cells_to_defined_names(self):
        model_compiler = ModelCompiler()
        archive = model_compiler.read_excel_file(
            testing.get_resource("reader.xlsm"))
        (
            model_compiler.model.cells,
            model_compiler.model.formulae,
            model_compiler.model.ranges
        ) = archive.read_cells(ignore_sheets=['Eleventh'])
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()
        model_compiler.link_cells_to_defined_names()

        self.assertEqual(self.model.cells, model_compiler.model.cells)

    @unittest.skip(
        'The ranges do not line up anymore until reader.json get regenerated. '
        'Shared formulas do not produce ranges anymore as they are resolved '
        'by the reader automatically.'
    )
    def test_build_ranges(self):
        model_compiler = ModelCompiler()
        archive = model_compiler.read_excel_file(
            testing.get_resource("reader.xlsm"))
        (
            model_compiler.model.cells,
            model_compiler.model.formulae,
            model_compiler.model.ranges
        ) = archive.read_cells(ignore_sheets=['Eleventh'])
        model_compiler.defined_names = archive.read_defined_names()
        model_compiler.build_defined_names()
        model_compiler.build_ranges()

        self.assertEqual(self.model.ranges, model_compiler.model.ranges)

    def test_extract_cells(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(
            testing.get_resource("reader.xlsm"), ignore_sheets=['Eleventh'])
        extracted_model = ModelCompiler.extract(
            reader_model, focus=['First!A2', 'First!B2', 'First!C2'])

        reference_model = Model()
        reference_model.set_cell_value('First!A2', 0.1)
        reference_model.set_cell_value('First!B2', 0.2)
        reference_model.set_cell_value('First!C2', 0.3)

        self.assertEqual(reference_model.cells, extracted_model.cells)

    def test_extract_defined_names(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(
            testing.get_resource("reader.xlsm"), ignore_sheets=['Eleventh'])
        extracted_model = ModelCompiler.extract(
            reader_model, focus=['Hundred', 'My_Range'])

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
        reference_model.defined_names['Hundred'] = deepcopy(
            reader_model.defined_names['Hundred'])
        reference_model.defined_names['My_Range'] = deepcopy(
            reader_model.defined_names['My_Range'])

        self.assertEqual(reference_model.cells, extracted_model.cells)
        self.assertEqual(
            reference_model.defined_names, extracted_model.defined_names)
        self.assertEqual(reference_model, extracted_model)

    def test_extract(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(
            testing.get_resource("reader.xlsm"), ignore_sheets=['Eleventh'])
        extracted_model = ModelCompiler.extract(
            reader_model,
            focus=['First!A2', 'First!B2', 'First!C2', 'Fourth!A2',
                   'Hundred', 'My_Range']
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
        reference_model.defined_names['Hundred'] = deepcopy(
            reader_model.defined_names['Hundred'])
        reference_model.defined_names['My_Range'] = deepcopy(
            reader_model.defined_names['My_Range'])
        reference_model.cells['Fourth!A2'] = reader_model.cells['Fourth!A2']

        self.assertEqual(reference_model.cells, extracted_model.cells)
        self.assertEqual(
            reference_model.defined_names,
            extracted_model.defined_names)

    def test_extract_and_evaluate(self):
        model_compiler = ModelCompiler()
        reader_model = model_compiler.read_and_parse_archive(
            testing.get_resource("model_compiler_and_evaluate.xlsx"))
        extracted_model = ModelCompiler.extract(
            reader_model,
            focus=['add_one']
        )

        self.assertEqual(1, extracted_model.get_cell_value('Sheet1!A1'))
        self.assertEqual(2, extracted_model.get_cell_value('Sheet1!B1'))
        extracted_model.set_cell_value('Sheet1!A1', 2)
        evaluator = Evaluator(extracted_model)
        evaluator.evaluate('add_one')
        self.assertEqual(3, extracted_model.get_cell_value('Sheet1!B1'))

    def test_dict_read_and_parse(self):

        input_dict = {
            "B4": 0.95,
            "B2": 1000,
            "B19": 0.001,
            "B20": 4,
            # B21
            "B22": 1,
            "B23": 2,
            "B24": 3,
            "B25": "=B2*B4",
            "B26": 5,
            "B27": 6,
            "B28": "=B19*B20*B22",
            "C22": "=SUM(B22:B28)",
        }

        model_compiler = ModelCompiler()
        my_model = model_compiler.read_and_parse_dict(input_dict)
        evaluator = Evaluator(my_model)

        # cells need a sheet and Sheet1 is default.
        evaluator.set_cell_value("Sheet1!B22", 100)
        self.assertEqual(0.4, evaluator.evaluate("Sheet1!B28"))
        self.assertEqual(1066.4, evaluator.evaluate("Sheet1!C22"))
