
import logging
from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile
import xml.etree.ElementTree as ET

from ..koala_types import XLCell
from ..koala_types import XLFormula


class Reader():
    """"""

    worksheet_metadata = {} # keyed on xl/workbook.xml {http://schemas.openxmlformats.org/officeDocument/2006/relationships}id
    defined_name_metadata = {} # keyed on xl/workbook.xml {http://schemas.openxmlformats.org/officeDocument/2006/relationships}definedNames
    shared_strings_metadata = {} # keyed on xl/sharedStrings.xml {http://schemas.openxmlformats.org/spreadsheetml/2006/main}si
    archive = None
    excel_file_name = None


    def __init__(self, file_name):
        self.excel_file_name = file_name


    def read(self):
        """Reads and parses the Excel file."""
        self.archive = Reader.build_archive(self.excel_file_name) # the file handle for the Excel file
        self.build_worksheet_metadata()
        self.build_defined_name_metadata()
        self.build_shared_string_metadata()


    def _parse_archive(self, archive_address):
        """Extracts data from a paricular part of the given Excel file."""

        self.build_archive(self.excel_file_name)
        try:
            return ET.fromstring( self.archive.read(archive_address) )
        except KeyError:
            return None



    @staticmethod
    def build_archive(file_name):
        """This creates the file handle for the given Excel file."""

        return ZipFile(file_name, 'r', ZIP_DEFLATED)


    def build_worksheet_metadata(self):
        """Extracts data about the worksheets to be found in this particular Excel file."""

        worksheet_root = self._parse_archive('xl/workbook.xml')

        if worksheet_root is not None:
            for sheet in worksheet_root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets"): # can get namespace from openpyxl but loading takes a long time
                self.worksheet_metadata[sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')] = {'name' : sheet.get('name'), 'sheetId' : sheet.get('sheetId')}

            relationship_root =  self._parse_archive('xl/_rels/workbook.xml.rels')

            for relationship in relationship_root:
                if relationship.get('Id') in self.worksheet_metadata.keys():
                    self.worksheet_metadata[relationship.get('Id')]['Target'] = relationship.get('Target')


    def build_defined_name_metadata(self):
        """Extracts data about the named cells in this particular Excel file."""

        defined_name_root = self._parse_archive('xl/workbook.xml')
        if defined_name_root is not None:
            defined_names = defined_name_root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedNames")
            if defined_names is not None and len(defined_names) > 0:
                for name in defined_names:
                    if name.get('hidden') is None:
                        self.defined_name_metadata[name.get('name')] = name.text


    def build_shared_string_metadata(self):
        """Extracts data about the shared strings in this particular Excel file.

        If a cell holds nothing but text, Excel places that text in a table and links the worksheet cell to it by a reference. 'for efficiency'.
        """

        shared_string_root = self._parse_archive('xl/sharedStrings.xml')
        if shared_string_root is not None:
            counter = 0
            for shared_string in shared_string_root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                sh_string = shared_string.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text
                sh_string = sh_string.replace('x005F_', '')
                self.shared_strings_metadata[counter] = sh_string
                counter += 1


    def read_cells(self, ignore_sheets=[], ignore_hidden=False):
        """Reads all cells from all non-ignored sheets and returns a dict of Cell objects keyed on the full cell address."""

        cells = {}
        formulae = {}
        ranges = {}
        for sheet_id in self.worksheet_metadata:

            sheet_name = self.worksheet_metadata[sheet_id]['name']
            if sheet_name not in ignore_sheets:

                logging.info( "Reading cells from {}".format(sheet_name) )
                worksheet_root = self._parse_archive("xl/%s" % self.worksheet_metadata[sheet_id]['Target'])

                rows = worksheet_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
                for row in rows:

                    columns = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                    for column in columns:

                        formula = column.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f')
                        if formula is not None:
                            if formula.get('t') is not None:
                                return_type = formula.get('t')
                            else:
                                return_type = "value"

                            # # Need to support shared formulas
                            # if formula.text is None:
                            #     print("FORMULA", formula.text, sheet_name, column.get('r'))
                            formula = XLFormula(formula.text, sheet_name=sheet_name, return_type=return_type, reference=formula.get('ref'))

                        value = column.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        if value is not None:
                            value = value.text

                        if 't' in column.attrib.keys():
                            # if column attribute 't' has value 's' there's only a text (string) in this cell
                            # we need to resolve the link to the shared string table and put the text as the value in our cell object
                            if column.attrib['t'] in ['s']:
                                value = self.shared_strings_metadata[int(value)]

                            # TODO: test a boolean value - could be legacy
                            elif column.attrib['t'] in ['b']:
                                value = self.shared_strings_metadata[bool(value)]

                            # TODO: test a numeric value - could be legacy
                            # if column attribute 't' has value 'n' there's a number in this cell
                            elif column.attrib['t'] in ['n']:
                                try:

                                    value = int(value)

                                except ValueError:
                                    value = float(value)

                            # if column attribute 't' has value 'array' we have found a dynamic array
                            elif column.attrib['t'] in ['array']:
                                # this cell is part of a dynamic array.
                                # we are not allowed to mutate this cell.
                                # we are likely to see a formula which defines the dynamic array
                                # TODO: support dynamic arrays
                                pass

                        else: # if the cell type is not specified there may be a number in the cell
                            try:

                                value = int(value)

                            except ValueError:
                                value = float(value)

                        cell_address = column.get('r')
                        should_eval = 'normal'

                        address = "{}!{}".format(sheet_name, cell_address)
                        cells[address] = XLCell(address, value = value, formula = formula)
                        if formula is not None:
                            formulae[address] = formula

        return [cells, formulae]


    def read_defined_names(self, ignore_sheets = [], ignore_hidden = False):
        """Reads all named cells from all non-ignored sheets and returns a dict of Cell objects keyed on the full cell address."""

        return self.defined_name_metadata
