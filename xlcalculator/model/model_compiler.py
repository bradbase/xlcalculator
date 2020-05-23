import logging
from copy import deepcopy

from ..read_excel import Reader
from ..xlcalculator_types import XLCell, XLFormula, XLRange
from .model import Model


class ModelCompiler:
    """Excel Workbook Data Model Compiler

    Factory class responsible for taking Microsoft Excel cells and named_range
    and create a model represented by a network graph that can be serialized
    to disk, and executed independently of Excel.
    """

    def __init__(self):
        self.model = Model()

    def read_excel_file(self, file_name):
        archive = Reader(file_name)
        archive.read()
        return archive

    def parse_archive(self, archive, ignore_sheets = [], ignore_hidden = False):
        self.model.cells, self.model.formulae, self.model.ranges = \
            archive.read_cells(ignore_sheets, ignore_hidden)
        self.defined_names = archive.read_defined_names(
            ignore_sheets, ignore_hidden)
        self.build_defined_names()
        self.link_cells_to_defined_names()
        self.build_ranges()

    def read_and_parse_archive(
            self, file_name=None, ignore_sheets = [], ignore_hidden = False,
            build_code=True
    ):
        archive = self.read_excel_file(file_name)
        self.parse_archive(archive, ignore_sheets=ignore_sheets)

        if build_code:
            self.model.build_code()

        return self.model

    def read_and_parse_dict(
            self, input_dict, default_sheet="Sheet1", build_code=True):
        for item in input_dict:
            if "!" in item:
                cell_address = item
            else:
                cell_address = "{}!{}".format(default_sheet, item)

            if (not Reader.isFloat(input_dict[item]) and
                     input_dict[item][0] == '='):
                cell = XLCell(
                    cell_address, None, formula=XLFormula(input_dict[item]))
                self.model.cells[cell_address] = cell
                self.model.formulae[cell_address] = cell.formula

            else:
                self.model.cells[cell_address] = XLCell(
                    cell_address, input_dict[item])

        self.build_ranges(default_sheet=default_sheet)

        if build_code:
            self.model.build_code()

        return self.model

    def build_defined_names(self):
        """Add defined ranges to model."""
        for name in self.defined_names:
            cell_address = self.defined_names[name]
            cell_address = cell_address.replace('$', '')

            # a cell has an address like; Sheet1!A1
            if ':' not in cell_address:
                if cell_address not in self.model.cells:
                    logging.warning(
                        f"Defined name {name} refers to empty cell "
                        f"{cell_address}. Is not being loaded.")
                    continue

                else:
                    if self.model.cells[cell_address] is not None:
                        self.model.defined_names[name] = \
                            self.model.cells[cell_address]

            else:
                self.model.defined_names[name] = XLRange(
                    cell_address, name=name)
                self.model.ranges[cell_address] = self.model.defined_names[name]

            if (cell_address in self.model.formulae and
                    name not in self.model.formulae
            ):
                self.model.formulae[name] = \
                    self.model.cells[cell_address].formula

    def link_cells_to_defined_names(self):
        for name in self.model.defined_names:
            defn = self.model.defined_names[name]

            if isinstance(defn, (XLCell)):
                self.model.cells[defn.address].defined_names.append(name)

            elif isinstance(defn, (XLRange)):
                if any(isinstance(el, list) for el in defn.cells):
                    for column in defn.cells:
                        for row_address in column:
                            self.model.cells[row_address].defined_names.append(
                                name)
                else:
                    # programmer error
                    message = "This isn't a dim2 array. {}".format(name)
                    logging.error(message)
                    raise Exception(message)
            else:
                message = (
                    f"Trying to link cells for {name}, but got unkown "
                    f"type {type(defn)}"
                )
                logging.error(message)
                raise ValueError(message)

    def build_ranges(self, default_sheet=None):
        for formula in self.model.formulae:
            associated_cells = set()
            for range in self.model.formulae[formula].terms:
                if "!" not in range:
                    range = "{}!{}".format(default_sheet, range)

                if ":" in range:
                    self.model.ranges[range] = XLRange(range, range)
                    associated_cells.update([
                        cell
                        for row in self.model.ranges[range].cells
                            for cell in row
                    ])
                else:
                    associated_cells.add( range )

                if range in self.model.ranges:
                    for row in self.model.ranges[range].cells:
                        for cell_address in row:
                            if cell_address not in self.model.cells.keys():
                                self.model.cells[cell_address] = XLCell(
                                    cell_address, '')

            if formula in self.model.cells:
                self.model.cells[formula].formula.associated_cells = \
                    associated_cells

            if formula in self.model.defined_names:
                self.model.defined_names[formula].formula.associated_cells = \
                    associated_cells

            self.model.formulae[formula].associated_cells = associated_cells

    @staticmethod
    def extract(model, focus):
        extracted_model = Model()

        for address in focus:
            if isinstance(address, str) and address in model.cells:
                extracted_model.cells[address] = deepcopy(model.cells[address])

            elif isinstance(address, str) and address in model.defined_names:

                extracted_model.defined_names[address] = defn = deepcopy(
                    model.defined_names[address])

                if isinstance(defn, XLCell):
                    extracted_model.cells[defn.address] = deepcopy(
                        model.cells[defn.address])

                elif isinstance(defn, XLRange):
                    for row in defn.cells:
                        for column in row:
                            extracted_model.cells[column] = deepcopy(
                                model.cells[column])

        for addr, cell in extracted_model.cells.items():
            if cell.formula is not None:
                for term in cell.formula.terms:
                    if (term in extracted_model.cells and
                            cell.formula != model.cells[addr].formula):
                        cell.formula = deepcopy(model.cells[addr].formula)

                    elif term not in extracted_model.cells:
                        extracted_model.cells[addr] = deepcopy(
                            model.cells[cell])

        extracted_model.build_code()

        return extracted_model
