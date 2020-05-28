import openpyxl

from . import patch, xltypes


class Reader():

    def __init__(self, file_name):
        self.excel_file_name = file_name

    def read(self):
        with patch.openpyxl_WorksheetReader_patch():
            self.book = openpyxl.load_workbook(self.excel_file_name)

    def read_defined_names(self, ignore_sheets=[], ignore_hidden=False):
        return {
            defn.name: defn.value
            for defn in self.book.defined_names.definedName
            if defn.hidden is None and defn.value != '#REF!'
        }

    def read_cells(self, ignore_sheets=[], ignore_hidden=False):
        cells = {}
        formulae = {}
        ranges = {}
        for sheet_name in self.book.sheetnames:
            if sheet_name in ignore_sheets:
                continue
            sheet = self.book[sheet_name]
            for cell in sheet._cells.values():
                addr = f'{sheet_name}!{cell.coordinate}'
                if cell.data_type == 'f':
                    formula = xltypes.XLFormula(cell.value, sheet_name)
                    formulae[addr] = formula
                    value = cell.cvalue
                else:
                    formula = None
                    value = cell.value

                cells[addr] = xltypes.XLCell(
                    addr, value=value, formula=formula)

        return [cells, formulae, ranges]
