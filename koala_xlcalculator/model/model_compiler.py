
import os.path
import logging
import re

import networkx as nx

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


    def parse_excel_file(self, file=None, ignore_sheets = [], ignore_hidden = False):
        """"""

        file_name = os.path.abspath(file)
        archive = Reader(file_name)
        archive.read()
        self.model.cells, self.model.formulae = archive.read_cells(ignore_sheets, ignore_hidden)
        self.defined_names = archive.read_defined_names(ignore_sheets, ignore_hidden)
        self.build_defined_names()
        self.link_cells_to_defined_names()
        self.build_ranges()

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
                # else:
                #     print("about to put a cell into the range dict", range)
                #     if range not in self.model.defined_names:
                #         self.model.ranges[range] = XLCell(range)


    def translate(self, outputs = [], inputs = []):
        """Translates a Microsoft Excel cell structure into a model representation."""

        self.model.translate(outputs=outputs, inputs=inputs)


    def persist(self, fname):
        """Convenience wrapper for the model persist_to_json_file."""

        self.model.persist_to_json_file(fname)
