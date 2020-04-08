
import os.path
import logging
import re

from ..read_excel import Reader
from ..koala_types import XLCell
from ..koala_types import XLRange
from .model import Model
from ..exceptions import ExcelError


class ModelCompiler():
    """Factory class responsible for taking Microsoft Excel cells and named_range and
    create a model represented by a network graph that can be serialized to disk,
    and executed independently of Excel.
    """

    def __init__(self):

        self.model = Model()
        self.defined_names = {}


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
        self.parse_archive( archive )
        return self.model


    def build_defined_names(self):
        """"""

        for name in self.defined_names:
            cell_address = self.defined_names[name]
            cell_address = cell_address.replace('$', '')

            # a cell has an address like; Sheet1!A1
            if ':' not in cell_address:
                self.model.defined_names[name] = XLCell(cell_address)

            # a range has an address like;
            # Sheet1!A1:A5
            # Sheet1!A1:E5
            # Sheet1!A1:A5,C1:C5,E1:E5
            else:
                self.model.defined_names[name] = XLRange(name, cell_address)
                self.model.ranges[cell_address] = self.model.defined_names[name]


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


    def build_ranges(self):
        """"""

        for formula in self.model.formulae:
            for range in self.model.formulae[formula].ranges:
                if ":" in range:
                    self.model.ranges[range] = XLRange(range, range)

                if range in self.model.ranges:
                    for r_ange in self.model.ranges[range].cells:
                        for cell_address in r_ange:
                            if cell_address not in self.model.cells.keys():
                                self.model.cells[cell_address] = XLCell(cell_address, '')
