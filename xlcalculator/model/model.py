import logging
from dataclasses import dataclass, field
import json
import gzip
from copy import copy

from jsonpickle import encode, decode

from .. import types
from ..read_excel import f_token
from . import parser


@dataclass
class Model():

    cells: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)
    formulae: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)
    ranges: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)
    defined_names: dict = field(
        init=False, default_factory=dict, compare=True, hash=True, repr=True)

    def set_cell_value(self, address, value):
        """Sets a new value for a specified cell."""
        if address in self.defined_names:
            if isinstance(self.defined_names[address], XLCell):
                address = self.defined_names[address].address

        if isinstance(address, str):
            if address in self.cells:
                self.cells[address].value = copy(value)
            else:
                self.cells[address] = types.XLCell(address, copy(value))

        elif isinstance(address, types.XLCell):
            if address.address in self.cells:
                self.cells[address.address].value = value
            else:
                self.cells[address.address] = XLCell(address.address, value)

        else:
            raise TypeError(
                f"Cannot set the cell value for an address of type "
                f"{address}. XLCell or a string is needed."
            )

    def get_cell_value(self, address):
        if address in self.defined_names:
            if isinstance(self.defined_names[address], XLCell):
                address = self.defined_names[address].address

        if isinstance(address, str):
            if address in self.cells:
                return self.cells[address].value
            else:
                logging.debug(
                    "Trying to get value for cell {address} but that cell "
                    "doesn't exist.")
                return 0

        elif isinstance(address, types.XLCell):
            if address.address in self.cells:
                return self.cells[address.address].value
            else:
                logging.debug(
                    "Trying to get value for cell {address.address} but "
                    "that cell doesn't exist")
                return 0

        else:
            raise TypeError(
                f"Cannot set the cell value for an address of type "
                f"{address}. XLCell or a string is needed."
            )

    def persist_to_json_file(self, fname):
        """Writes the state to disk.

        Doesn't write the graph directly, but persist all the things that
        provide the ability to re-create the graph.
        """
        output = {
            'cells' : self.cells,
            'defined_names' : self.defined_names,
            'formulae' : self.formulae,
            'ranges' : self.ranges,
        }

        if fname.split('.')[-1:][0].upper() in ['GZIP', 'GZ']:
            outfile = gzip.GzipFile(fname,'wb')
            outfile.write( str.encode( encode(output, keys=True) ) )
        else:
            outfile = open(fname, "w")
            outfile.write( encode(output, keys=True) )

        outfile.close()

    def construct_from_json_file(self, fname, build_code=False):
        """Constructs a graph from a state persisted to disk."""

        if fname.split('.')[-1:][0].upper() in ['GZIP', 'GZ']:
            infile = gzip.GzipFile(fname,'rb')
        else:
            infile = open(fname, "rb")

        json_bytes = infile.read()
        infile.close()
        data = decode(
            json_bytes, keys=True,
            classes=(types.XLCell, types.XLFormula, f_token, types.XLRange))
        self.cells = data['cells']

        self.defined_names = data['defined_names']
        self.ranges = data['ranges']
        self.formulae = data['formulae']

        if build_code:
            self.build_code()

    def build_code(self):
        """Define the Python code for all cells in the dict of cells."""

        for cell in self.cells:
            if self.cells[cell].formula is not None:
                defined_names = {
                    name: defn.address
                    for name, defn in self.defined_names.items()}
                self.cells[cell].formula.ast = parser.FormulaParser().parse(
                    self.cells[cell].formula.formula, defined_names)

    def __eq__(self, other):

        cells_comparison = []
        for self_cell in self.cells:
            cells_comparison.append(
                self.cells[self_cell] == other.cells[self_cell])

        defined_names_comparison = []
        for self_defined_names in self.defined_names:
            defined_names_comparison.append(
                self.defined_names[self_defined_names]
                    == other.defined_names[self_defined_names])

        return (
            self.__class__ == other.__class__ and
            all(cells_comparison) and
            all(defined_names_comparison)
        )
