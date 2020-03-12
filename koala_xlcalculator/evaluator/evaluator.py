
import logging
from copy import deepcopy

import pandas as pd

from ..exceptions import ExcelError
from ..koala_types import XLCell, XLRange
from ..function_library import xSum


class Evaluator():
    """Traverses and evaluates a given model."""

    def __init__(self, model):

        self.model = model


    def evaluate(self, cell, is_addr=True):
        """Evaluates Python code as defined in formula.python_code"""

        if is_addr:
            try:

                # Although defined names have been resolved in Model.create_node()
                # we need to attempt to resolve defined names as we might have been
                # given one in argument cell.
                if cell in self.model.defined_names:
                    name_definition = self.model.defined_names[cell]
                    if isinstance(name_definition, XLCell):
                        cell = self.model.cells[name_definition.address]

                    elif isinstance(name_definition, XLRange):
                        message = "I can't resolve {} to a cell. It's a range and they aren't supported yet.".format(cell)
                        logging.error(message)
                        raise Exception(message)

                    elif isinstance(name_definition, XLFormula):
                        message = "I can't resolve {} to a cell. It's a formula and they aren't supported as a cell reference.".format(cell)
                        logging.error(message)
                        raise Exception(message)

                elif cell in self.model.ranges:

                    cell = self.model.cells[cell]


                else:
                    cell = self.model.cells[cell]

            except:
                logging.error('Empty cell at {}'.format(cell))
                return ExcelError('#NULL', 'Cell {} is empty'.format(cell))

        # no formula, or no evaluation means fixed value
        if cell.formula is None or cell.formula.evaluate == False:
            logging.debug("\r\nCell {} has no formula \r\nbut its value is {}".format(cell.address, cell.value))
            return cell.value

        try:
            if cell.formula.python_code != None:
                value = eval(cell.formula.python_code)
                if isinstance(value, Evaluator): # this should mean that vv is the result of RangeCore.apply_all, but with only one value inside
                    cell.value = value.values[0]

                else:
                    cell.value = value if value != '' else None

            else:
                cell.value = 0

            cell.need_update = False

        except Exception as exception:
            if str(exception).startswith("Problem evalling"):
                raise exception

            else:
                raise Exception("Problem evalling: {} for {}, {}".format(exception, cell.address, cell.formula.python_code))

        logging.debug("\r\nCell {} has a formula, {} \r\nwhich has been translated to Python as {} which evaluates to {}\r\n".format(cell.address, cell.formula.formula, cell.formula.python_code, cell.value))

        return cell.value


    def eval_ref(self, address):
        """"""

        if address in self.model.cells:
            return pd.DataFrame([self.model.cells[address]])

        elif address in self.model.defined_names:
            return pd.DataFrame([self.model.defined_names[address]])

        elif address in self.model.ranges:
            range_cells = []
            for range_column in self.model.ranges[address].cells:
                row = []

                for cell_address in range_column:
                    row.append( self.model.cells[cell_address].value )

                range_cells.append(row)

            self.model.ranges[address].value = pd.DataFrame(range_cells)

            return self.model.ranges[address]


    @staticmethod
    def apply(func, first, second, ref=None):
        """"""

        # TODO: need to support ranges
        return Evaluator.apply_all(func, first, second)


    @staticmethod
    def apply_one(func, first, second, ref=None):
        """"""

        #TODO: This can't be called ATM, but needs to be - once we fully support ranges.

        function = SUPPORTED_OPERATORS[func]

        if ref is None:
            first_value = first
            second_value = second

        else:
            first_value = self.find_associated_value(ref, first)
            second_value = self.find_associated_value(ref, second)

        return function(first_value, second_value)


    @staticmethod
    def apply_all(func, first, second, ref=None):
        """"""

        #TODO: currently inadequate as Ranges are not correctly or fully supported.

        function = SUPPORTED_OPERATORS[func]

        if isinstance(first, XLRange) and isinstance(second, XLRange):
            if first.length != second.length:
                raise ExcelError('#VALUE!', 'apply_all must have 2 Ranges of identical length')

            vals = [function(
                x.value if isinstance(x, CellBase) else x,
                y.value if isinstance(x, CellBase) else y
            ) for x, y in zip(first.cells, second.cells)]

            return (first.addresses, vals, first.nrows, first.ncols)

        elif isinstance(first, XLRange):
            vals = [function(
                x.value if isinstance(x, CellBase) else x,
                second
            ) for x in first.cells]

            return (first.addresses, vals, first.nrows, first.ncols)

        elif isinstance(second, XLRange):
            vals = [function(
                first,
                x.value if isinstance(x, CellBase) else x
            ) for x in second.cells]

            return (second.addresses, vals, second.nrows, second.ncols)

        else:
            return function(first, second)


    @staticmethod
    def find_associated_cell(ref, range):
        """This function retrieves the cell associated to ref in a Range
        For instance, in the range [A1, B1, C1], the cell associated to B2 is B1
        This is useful to mimic the way Excel works.
        """

        if ref is not None:
            row, col = ref

            if (range.length) == 0:  # if a Range is empty, it means normally that all its cells are empty
                return None

            elif range.type == "vertical":
                if (row, range.origin[1]) in range.order:
                    return range.addresses[
                        range.order.index((row, range.origin[1]))
                    ]

                else:
                    return None

            elif range.type == "horizontal":
                if (range.origin[0], col) in range.order:
                    return range.addresses[
                        range.order.index((range.origin[0], col))
                    ]

                else:
                    return None

            elif range.type == "scalar":
                if (row, range.origin[1]) in range.order:
                    return range.addresses[
                        range.order.index((row, range.origin[1]))
                    ]

                elif (range.origin[0], col) in range.order:
                    return range.addresses[
                        range.order.index((range.origin[0], col))
                    ]

                elif (row, col) in range.order:
                    return range.addresses[
                        range.order.index((row, col))
                    ]

                else:
                    return None

            else:
                return None

        else:
            return None


    @staticmethod
    def find_associated_value(ref, item):
        """retrieves the value and not the Cell."""

        row, col = ref

        if isinstance(item, Evaluator):
            try:
                if (item.length) == 0:  # if a Range is empty, it means normally that all its cells are empty
                    item_value = 0

                elif item.type == "vertical":
                    if item.__cellmap is not None:
                        item_value = item[(row, item.origin[1])].value if (row, item.origin[1]) in item else None

                    else:
                        item_value = item[(row, item.origin[1])] if (row, item.origin[1]) in item else None

                elif item.type == "horizontal":

                    if item.__cellmap is not None:
                        try:

                            item_value = item[(item.origin[0], col)].value if (item.origin[0], col) in item else None

                        except:
                            raise Exception

                    else:
                        item_value = item[(item.origin[0], col)] if (item.origin[0], col) in item else None

                else:
                    raise ExcelError('#VALUE!', 'cannot use find_associated_value on %s' % item.type)

            except ExcelError as e:
                raise Exception('First argument of Range operation is not valid: ' + e.value)

        elif item is None:
            item_value = 0

        else:
            item_value = item

        return item_value


    @staticmethod
    def check_value(a):
        """"""

        if isinstance(a, ExcelError):
            return a

        elif isinstance(a, str) and a in ErrorCodes:
            return ExcelError(a)

        try:  # This is to avoid None or Exception returned by Range operations
            if float(a) or isinstance(a, (str)):
                return a

            else:
                return 0

        except:
            if a == 'True':
                return True

            elif a == 'False':
                return False

            else:
                return 0


    @staticmethod
    def add(a, b):
        """"""

        try:
            return Evaluator.check_value(a.value) + Evaluator.check_value(b.value)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def substract(a, b):
        """"""

        try:
            return Evaluator.check_value(a.value) - Evaluator.check_value(b.value)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def minus(a, b=None):
        """"""

        # b is not used, but needed in the signature. Maybe could be better
        try:
            return -Evaluator.check_value(a.value)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def multiply(a, b):
        """"""

        try:
            return Evaluator.check_value(a.value) * Evaluator.check_value(b.value)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def divide(a, b):
        """"""

        try:
            return float(Evaluator.check_value(a.value)) / float(Evaluator.check_value(b.value))

        except Exception as e:
            return ExcelError('#DIV/0!', e)


    @staticmethod
    def is_strictly_superior(a, b):
        try:
            return Evaluator.check_value(a.value) > Evaluator.check_value(b.value)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_strictly_inferior(a, b):
        try:
            return Evaluator.check_value(a.value) < Evaluator.check_value(b.value)
        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_superior_or_equal(a, b):
        try:
            a = Evaluator.check_value(a.value)
            b = Evaluator.check_value(b.value)

            return a > b or Evaluator.is_almost_equal(a, b)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_inferior_or_equal(a, b):
        try:
            a = Evaluator.check_value(a.value)
            b = Evaluator.check_value(b.value)

            return a < b or Evaluator.is_almost_equal(a, b)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_equal(a, b):
        """"""

        try:
            if not isinstance(a, (str)):
                a = Evaluator.check_value(a)

            if not isinstance(b, (str)):
                b = Evaluator.check_value(b)

            return Evaluator.is_almost_equal(a, b, precision=0.00001)

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_not_equal(a, b):
        """"""

        try:
            if not isinstance(a, (str)):
                a = Evaluator.check_value(a)

            if not isinstance(a, (str)):
                b = Evaluator.check_value(b)

            return a != b

        except Exception as e:
            return ExcelError('#N/A', e)


    @staticmethod
    def is_number(s): # http://stackoverflow.com/questions/354038/how-do-i-check-if-a-string-is-a-number-float-in-python
        try:
            float(s)
            return True

        except:
            return False


    @staticmethod
    def is_almost_equal(a, b, precision = 0.0001):
        if Evaluator.is_number(a) and Evaluator.is_number(b):
            return abs(float(a) - float(b)) <= precision

        elif (a is None or a == 'None') and (b is None or b == 'None'):
            return True

        else: # booleans or strings
            return str(a) == str(b)


SUPPORTED_OPERATORS = {
    "minus": Evaluator.minus,
    "add": Evaluator.add,
    "substract": Evaluator.substract,
    "multiply": Evaluator.multiply,
    "divide": Evaluator.divide,
    "is_equal": Evaluator.is_equal,
    "is_not_equal": Evaluator.is_not_equal,
    "is_strictly_superior": Evaluator.is_strictly_superior,
    "is_strictly_inferior": Evaluator.is_strictly_inferior,
    "is_superior_or_equal": Evaluator.is_superior_or_equal,
    "is_inferior_or_equal": Evaluator.is_inferior_or_equal,
}
