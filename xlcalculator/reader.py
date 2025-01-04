from os import PathLike
from typing import Union
import openpyxl

from . import patch, xltypes

# TODO: remove unused variables


class Reader:
    def __init__(self, file_name: Union[str, PathLike]):
        self.excel_file_name = file_name

    def read(self) -> None:
        with patch.openpyxl_WorksheetReader_patch():
            self.book = openpyxl.load_workbook(self.excel_file_name)

    def read_defined_names(self, ignore_sheets=[], ignore_hidden=False) -> dict:
        return {
            defn.name: defn.value
            for name, defn in self.book.defined_names.items()
            if defn.hidden is None and defn.value != "#REF!"
        }

    def read_cells(
        self, ignore_sheets: list[str] = [], ignore_hidden=False
    ) -> tuple[dict[str, xltypes.XLCell], dict[str, xltypes.XLFormula], set]:
        cells: dict[str, xltypes.XLCell] = {}
        formulae: dict[str, xltypes.XLFormula] = {}
        ranges: set = {}  # TODO: what is this for?
        for sheet_name in self.book.sheetnames:
            if sheet_name in ignore_sheets:
                continue
            sheet = self.book[sheet_name]
            for cell in sheet._cells.values():
                addr = f"{sheet_name}!{cell.coordinate}"
                if cell.data_type == "f":
                    value = cell.value
                    if isinstance(value, openpyxl.worksheet.formula.ArrayFormula):
                        value = value.text
                    formula = xltypes.XLFormula(value, sheet_name)
                    formulae[addr] = formula
                    value = cell.cvalue
                else:
                    formula = None
                    value = cell.value

                cells[addr] = xltypes.XLCell(addr, value=value, formula=formula)

        return [cells, formulae, ranges]
