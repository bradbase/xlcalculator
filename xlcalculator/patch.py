import contextlib
import mock
import openpyxl


class Cell(openpyxl.cell.cell.Cell):

    __slots__ = openpyxl.cell.cell.Cell.__slots__ + (
        'cvalue',
    )


class WorkSheetParser(openpyxl.worksheet._reader.WorkSheetParser):

    def parse_cell(self, element):
        cell = super().parse_cell(element)
        if cell['data_type'] == 'f':
            with mock.patch.object(self, 'data_only', True):
                cell_data = super().parse_cell(element)
            cell['cvalue'] = cell_data.get('value')
        return cell


class WorksheetReader(openpyxl.worksheet._reader.WorksheetReader):

    def __init__(self, ws, xml_source, shared_strings, data_only):
        super().__init__(ws, xml_source, shared_strings, data_only)
        self.parser = WorkSheetParser(
            xml_source, shared_strings, data_only, ws.parent.epoch,
            ws.parent._date_formats)

    def bind_cells(self):
        for idx, row in self.parser.parse():
            for cell in row:
                style = self.ws.parent._cell_styles[cell['style_id']]
                c = Cell(
                    self.ws, row=cell['row'], column=cell['column'],
                    style_array=style)
                c._value = cell['value']
                c.data_type = cell['data_type']
                # xlcalculator extension: Store the cached value as well.
                if c.data_type == 'f':
                    c.cvalue = cell['cvalue']
                self.ws._cells[(cell['row'], cell['column'])] = c
        self.ws.formula_attributes = self.parser.array_formulae
        if self.ws._cells:
            self.ws._current_row = self.ws.max_row


@contextlib.contextmanager
def openpyxl_WorksheetReader_patch():
    """Allow extraction of cached formula value."""
    with mock.patch.object(
            openpyxl.worksheet._reader, 'WorksheetReader', WorksheetReader):
        with mock.patch.object(
                openpyxl.reader.excel, 'WorksheetReader', WorksheetReader):
            yield
