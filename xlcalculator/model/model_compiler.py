
import os.path
import logging
import re
from copy import deepcopy

from xlfunctions.exceptions import ExcelError

from ..read_excel import Reader
from ..xlcalculator_types import XLCell
from ..xlcalculator_types import XLRange
from ..xlcalculator_types import XLFormula
from .model import Model


class ModelCompiler():
    """Factory class responsible for taking Microsoft Excel cells and named_range and
    create a model represented by a network graph that can be serialized to disk,
    and executed independently of Excel.
    """

    def __init__(self):

        self.model = Model()


    @staticmethod
    def read_excel_file(file):
        """"""
        file_name = os.path.abspath(file)
        archive = Reader(file_name)
        archive.read()
        return archive


    def parse_archive(self, archive, ignore_sheets = [], ignore_hidden = False):
        """"""
        self.model.cells, self.model.formulae, self.model.ranges = archive.read_cells(ignore_sheets, ignore_hidden)
        self.defined_names = archive.read_defined_names(ignore_sheets, ignore_hidden)
        self.build_defined_names()
        self.link_cells_to_defined_names()
        self.build_ranges()


    def read_and_parse_archive(self, file_name=None, ignore_sheets = [], ignore_hidden = False):
        """"""
        archive = ModelCompiler.read_excel_file(file_name)
        self.parse_archive(archive, ignore_sheets=ignore_sheets)
        return deepcopy(self.model)


    def read_and_parse_dict(self, input_dict, default_sheet="Sheet1"):
        """"""
        for item in input_dict:
            if "!" in item:
                cell_address = item
            else:
                cell_address = "{}!{}".format(default_sheet, item)

            if not Reader.isFloat(input_dict[item]) and input_dict[item][0] == '=':
                self.model.cells[cell_address] = XLCell(cell_address, None, formula=XLFormula(input_dict[item]))
                self.model.formulae[cell_address] = self.model.cells[cell_address].formula

            else:
                self.model.cells[cell_address] = XLCell(cell_address, input_dict[item])

        self.build_ranges(default_sheet=default_sheet)

        return deepcopy(self.model)


    def build_defined_names(self):
        """"""

        for name in self.defined_names:
            cell_address = self.defined_names[name]
            cell_address = cell_address.replace('$', '')

            # a cell has an address like; Sheet1!A1
            if ':' not in cell_address:
                self.model.defined_names[name] = self.model.cells[cell_address]

            # a range has an address like;
            # Sheet1!A1:A5
            # Sheet1!A1:E5
            # Sheet1!A1:A5,C1:C5,E1:E5
            else:
                self.model.defined_names[name] = XLRange(name, cell_address)
                self.model.ranges[cell_address] = self.model.defined_names[name]

            if cell_address in self.model.formulae and name not in self.model.formulae:
                self.model.formulae[name] = self.model.cells[cell_address].formula


    def link_cells_to_defined_names(self):
        """"""

        for name in self.model.defined_names:
            definition = self.model.defined_names[name]

            if isinstance(definition, (XLCell)):
                self.model.cells[definition.address].defined_names.append(name)

            elif isinstance(definition, (XLRange)):
                if any(isinstance(el, list) for el in definition.cells):
                    for column in definition.cells:
                        for row_address in column:
                            self.model.cells[row_address].defined_names.append(name)
                else:
                    # programmer error
                    message = "This isn't a dim2 array. {}".format(name)
                    logging.error(message)
                    raise Exception(message)
            else:
                message = "Trying to link cells for {} and {} is not recognisable {}".format(name, type(definition))
                logging.error(message)
                raise Exception(message)


    def build_ranges(self, default_sheet=None):
        """"""

        for formula in self.model.formulae:
            associated_cells = set()
            for range in self.model.formulae[formula].terms:
                if "!" not in range:
                    range = "{}!{}".format(default_sheet, range)

                if ":" in range:
                    self.model.ranges[range] = XLRange(range, range)
                    associated_cells.update( [cell for row in self.model.ranges[range].cells for cell in row] )

                else:
                    associated_cells.add( range )

                if range in self.model.ranges:
                    for row in self.model.ranges[range].cells:
                        for cell_address in row:
                            if cell_address not in self.model.cells.keys():
                                self.model.cells[cell_address] = XLCell(cell_address, '')

            if formula in self.model.cells:
                self.model.cells[formula].formula.associated_cells = associated_cells

            if formula in self.model.defined_names:
                self.model.defined_names[formula].formula.associated_cells = associated_cells

            self.model.formulae[formula].associated_cells = associated_cells


    @staticmethod
    def extract(model, focus):
        extracted_model = Model()

        for address in focus:
            if isinstance(address, str) and address in model.cells:
                extracted_model.cells[address] = deepcopy(model.cells[address])

            elif isinstance(address, str) and address in model.defined_names:
                extracted_model.defined_names[address] = deepcopy(model.defined_names[address])
                if isinstance(extracted_model.defined_names[address], XLCell):
                    extracted_model.cells[model.defined_names[address].address] = deepcopy(model.cells[ model.defined_names[address].address ])

                elif isinstance(extracted_model.defined_names[address], XLRange):
                    for row in extracted_model.defined_names[address].cells:
                        for column in row:
                            extracted_model.cells[column] = deepcopy(model.cells[column])

        for cell in extracted_model.cells:
            if extracted_model.cells[cell].formula is not None:
                for term in extracted_model.cells[cell].formula.terms:
                    if term in extracted_model.cells and extracted_model.cells[cell].formula != model.cells[cell].formula:
                        extracted_model.cells[cell].formula = deepcopy(model.cells[cell].formula)

                    elif term not in extracted_model.cells:
                        extracted_model.cells[cell] = deepcopy(model.cells[cell])

        extracted_model.build_code()

        return extracted_model
