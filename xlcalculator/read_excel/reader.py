
import logging
from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile
import xml.etree.cElementTree as ET
from decimal import Decimal
from decimal import Context
from decimal import ROUND_FLOOR
import re
from copy import copy
from datetime import datetime
from math import e

from ..xlcalculator_types import XLCell
from ..xlcalculator_types import XLFormula
from ..xlcalculator_types import XLRange


class Reader():
    """"""

    precision_mask = "{0:." + str(17) + "f}1"
    precision = precision_mask.format(e * 10**-17)

    def __init__(self, file_name):
        self.archive = None
        self.excel_file_name = file_name
        self.worksheet_metadata = {} # keyed on xl/workbook.xml {http://schemas.openxmlformats.org/officeDocument/2006/relationships}id
        self.defined_name_metadata = {} # keyed on xl/workbook.xml {http://schemas.openxmlformats.org/officeDocument/2006/relationships}definedNames
        self.shared_strings_metadata = {} # keyed on xl/sharedStrings.xml {http://schemas.openxmlformats.org/spreadsheetml/2006/main}si


    @staticmethod
    def isInteger(value):
        value = str(value)
        return value=='0' or (value if value.find('..') > -1 else value.lstrip('-+').rstrip('0').rstrip('.')).isdigit()


    @staticmethod
    def isFloat(value):
        if value is None:
            return False

        try:
            float(value)
            return True
        except ValueError:
            return False


    def read(self):
        """Reads and parses the Excel file."""
        self.archive = Reader.build_archive(self.excel_file_name) # the file handle for the Excel file
        self.build_worksheet_metadata()
        # self.build_shared_string_metadata()


    def _parse_archive(self, archive_address):
        """Extracts data from a paricular part of the given Excel file."""

        if archive_address[:1] == '/': # xlsx made by pyopenxl sometimes puts leadng /
            archive_address = archive_address[1:]

        self.archive = self.build_archive(self.excel_file_name)
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

        # get an iterable
        for fname in ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/sharedStrings.xml', '[Content_Types].xml']:
            try:
                with self.archive.open(fname) as f:
                    context = ET.iterparse(f, events=("start", "end"))
                    # turn it into an iterator
                    # context = iter(context)
                    # get the root element
                    event, root = context.__next__()

                    if root.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbook':
                        for event, elem in context:
                            if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName':
                                if elem.get('hidden') is None and elem.text not in ['#REF!']:
                                    self.defined_name_metadata[elem.get('name')] = elem.text

                            if elem.tag in ['{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet']:
                                self.worksheet_metadata[elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')] = {'name' : elem.get('name'), 'sheetId' : elem.get('sheetId')}

                            if event == "end" and elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName':
                                root.clear()

                    if root.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sst':
                        counter = 0
                        for event, elem in context:
                            if event == 'start' and elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t':
                                sh_string = elem.text
                                sh_string = sh_string.replace('x005F_', '')
                                self.shared_strings_metadata[counter] = sh_string
                                counter += 1


                            if event == "end" and elem.tag == "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sst":
                                root.clear()

                    # if root.tag == '{http://schemas.openxmlformats.org/package/2006/content-types}Types':
                    #     for event, elem in context:
                    #         if event == 'start' and elem.get('ContentType') == 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml':
                    #             if "rId{}".format(str(counter)) in self.worksheet_metadata:
                    #                 self.worksheet_metadata[sheet_name] = copy(self.worksheet_metadata["rId{}".format(str(counter))])
                    #
                    #         if event == "end" and elem.tag == "{http://schemas.openxmlformats.org/package/2006/content-types}Override":
                    #             root.clear()

                    if root.tag == '{http://schemas.openxmlformats.org/package/2006/relationships}Relationships':
                        for event, elem in context:
                            if event == 'start' and elem.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':
                                self.worksheet_metadata[elem.get('Id')]['Target'] = elem.get('Target')

            except:
                logging.info( "File {} is not in archive.".format(fname) )


    # def build_shared_string_metadata(self):
    #     """Extracts data about the shared strings in this particular Excel file.
    #
    #     If a cell holds nothing but text, Excel places that text in a table and links the worksheet cell to it by a reference. 'for efficiency'.
    #     """
    #
    #     shared_string_root = self._parse_archive('xl/sharedStrings.xml')
    #     if shared_string_root is not None:
    #         counter = 0
    #         for shared_string in shared_string_root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
    #             the_string = shared_string.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
    #             if the_string is not None:
    #                 sh_string = the_string.text
    #                 sh_string = sh_string.replace('x005F_', '')
    #                 self.shared_strings_metadata[counter] = sh_string
    #                 counter += 1


    def read_cells(self, ignore_sheets=[], ignore_hidden=False):
        """Reads all cells from all non-ignored sheets and returns a dict of Cell objects keyed on the full cell address."""

        cells = {}
        formulae = {}
        ranges = {}
        # shared_formulae = {}
        # shared_formula_offset = 0
        # for sheet_id in self.worksheet_metadata:
        #     sheet_name = self.worksheet_metadata[sheet_id]['name']
        #     if sheet_name not in ignore_sheets:
        #
        #         logging.info( "Reading cells from {}".format(sheet_name) )
        #         worksheet_root = self._parse_archive("%s" % self.worksheet_metadata[sheet_id]['Target'])
        #
        #         decimal_context = Context(prec=15, rounding=ROUND_FLOOR)
        #
        #         rows = worksheet_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        #         for row in rows:
        #
        #             columns = row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
        #             for column in columns:
        #
        #                 formula = column.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f')
        #                 if formula is not None:
        #                     shared_index = None
        #                     shared_formula = None
        #                     shared_formula_range = None
        #
        #                     if formula.get('t') is not None:
        #                         return_type = formula.get('t')
        #
        #                         if return_type == 'shared':
        #                             shared_index = formula.get('si')
        #                             shared_formula = column.get('r')
        #                             shared_formula_offset += 1
        #                             if formula.text is not None:
        #                                 shared_formulae[formula.get('si')] = {'formula' : formula.text, 'ref' : formula.get('ref')}
        #                                 shared_formula_offset = 0
        #
        #                         if return_type == 'array' and ":" in formula.get('ref'):
        #                             range_address = "{}!{}".format(sheet_name, formula.get('ref'))
        #                             ranges[range_address] = XLRange(range_address, range_address, formula=formula.text)
        #
        #                     else:
        #                         return_type = "value"
        #
        #                     if  formula.get('t') == 'shared' and formula.text is not None:
        #                         formula = XLFormula(formula.text, sheet_name=sheet_name, return_type=return_type, reference=formula.get('ref'), shared_formula_offset=shared_formula_offset, shared_formula_range=shared_formulae[shared_index]['ref'])
        #
        #                     elif formula.get('t') == 'shared' and formula.text is None:
        #                         formula = XLFormula(shared_formulae[shared_index]['formula'], sheet_name=sheet_name, return_type=return_type, reference=formula.get('ref'), shared_formula_offset=shared_formula_offset, shared_formula_range=shared_formulae[shared_index]['ref'])
        #
        #                     else:
        #                         formula = XLFormula(formula.text, sheet_name=sheet_name, return_type=return_type, reference=formula.get('ref'))
        #
        #                 value = column.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        #                 if value is not None:
        #                     value = value.text
        #
        #                     if 't' in column.attrib.keys():
        #                         # if column attribute 't' has value 's' there's only a text (string) in this cell
        #                         # we need to resolve the link to the shared string table and put the text as the value in our cell object
        #                         if column.attrib['t'] in ['s']:
        #                             # unicode strings don't work in the shared strings mechanism.
        #                             value = self.shared_strings_metadata[int(value)]
        #
        #                         # TODO: test a boolean value - could be legacy
        #                         elif column.attrib['t'] in ['b']:
        #                             value = self.shared_strings_metadata[bool(value)]
        #
        #                         # TODO: test a numeric value - could be legacy
        #                         # if column attribute 't' has value 'n' there's a number in this cell
        #                         elif column.attrib['t'] in ['n']:
        #
        #                             if Reader.isInteger(value):
        #                                 value = int(value)
        #
        #                             elif Reader.isFloat(value):
        #                                 decimal_from_float = decimal_context.create_decimal_from_float(float(value))
        #                                 decimal_value = Decimal(value)
        #                                 diff = decimal_value - decimal_from_float
        #                                 value = float(decimal_from_float + diff)
        #
        #                         # if column attribute 't' has value 'array' we have found a dynamic array
        #                         elif column.attrib['t'] in ['array']:
        #                             # this cell is part of a dynamic array.
        #                             # we are not allowed to mutate this cell.
        #                             # we are likely to see a formula which defines the dynamic array
        #                             # TODO: support dynamic arrays
        #                             pass
        #
        #                     else: # if the cell type is not specified there may be a number in the cell
        #                         if Reader.isInteger(value):
        #                             value = int(value)
        #
        #                         elif Reader.isFloat(value):
        #                             decimal_from_float = decimal_context.create_decimal_from_float(float(value))
        #                             decimal_value = Decimal(value)
        #                             diff = decimal_value - decimal_from_float
        #                             value = float(decimal_from_float + diff)
        #
        #                 cell_address = column.get('r')
        #                 should_eval = 'normal'
        #
        #                 address = "{}!{}".format(sheet_name, cell_address)
        #                 cells[address] = XLCell(address, value = value, formula = formula)
        #
        #                 if formula is not None:
        #                     formulae[address] = formula

        start_time = datetime.now()
        decimal_context = Context(prec=15, rounding=ROUND_FLOOR)

        worksheets = []
        for item in self.worksheet_metadata:
            if self.worksheet_metadata[item]['name'] not in ignore_sheets:
                if self.worksheet_metadata[item]['Target'][0] == '/':
                    worksheets.append( (self.worksheet_metadata[item]['name'], self.worksheet_metadata[item]['Target'][1:]) )
                elif self.worksheet_metadata[item]['Target'][0] == 'w':
                    worksheets.append( (self.worksheet_metadata[item]['name'], "xl/{}".format(self.worksheet_metadata[item]['Target'])) )

        for sheet_name, fname in worksheets:
            with self.archive.open(fname) as f:
                context = ET.iterparse(f, events=("start", "end"))
                # turn it into an iterator
                # context = iter(context)
                # get the root element
                event, root = context.__next__()

                if root.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}worksheet':
                    cellparts = {'address' : None, 'value': None, 'formula': None, 'ref' : None}
                    rowcounter = 0
                    cellcounter = 0
                    time_before = datetime.now()
                    is_shared_string = False
                    for event, elem in context:

                        if event == 'start':

                            if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c':
                                cellparts['address'] = elem.get('r')
                                if elem.get('t') == 's':
                                    is_shared_string = True

                            if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v':
                                if is_shared_string:
                                    cellparts['value'] = self.shared_strings_metadata[int(elem.text)]

                                elif Reader.isFloat(elem.text):
                                    decimal_from_float = decimal_context.create_decimal_from_float(float(elem.text))
                                    decimal_value = Decimal(elem.text)
                                    diff = decimal_value - decimal_from_float
                                    cellparts['value'] = float(decimal_from_float + diff)

                                elif Reader.isInteger(elem.text):
                                    cellparts['value'] = int(elem.text)

                                else:
                                    # not sure how (due to shared strings), but this is assumed to be a string.
                                    cellparts['value'] = elem.text

                            if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f':
                                cellparts['formula'] = elem.text

                                if elem.get('t') == 'array':
                                    cellparts['ref'] = elem.get('ref')

                        if event == "end" and elem.tag == "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c":
                            address = "{}!{}".format(sheet_name, cellparts['address'])
                            if cellparts['formula'] is not None:
                                cells[address] = XLCell(address, value = cellparts['value'], formula = XLFormula(cellparts['formula'], sheet_name=sheet_name))

                                formulae[address] = cells[address].formula
                                if cellparts['ref'] is not None and ":" in cellparts['ref']:
                                    range_address = "{}!{}".format(sheet_name, cellparts['ref'])
                                    ranges[range_address] = XLRange(range_address, range_address)

                            else:
                                cells[address] = XLCell(address, value = cellparts['value'])
                            cellcounter += 1
                            is_shared_string = False


                        if event == "end" and elem.tag == "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row":
                            cellparts = {'address' : None, 'value': None, 'formula': None, 'ref' : None}
                            rowcounter += 1
                            # if rowcounter % 10000 == 0 and cellcounter == rowcounter * 26:
                            #     time_now = datetime.now()
                            #     print("done {} rows. cell count {}, {} {}".format(rowcounter, cellcounter, time_now - time_before, time_now - start_time))
                            #     time_before = time_now
                            root.clear()

        return [cells, formulae, ranges]


    def read_defined_names(self, ignore_sheets = [], ignore_hidden = False):
        """Reads all named cells from all non-ignored sheets and returns a dict of Cell objects keyed on the full cell address."""

        return self.defined_name_metadata
