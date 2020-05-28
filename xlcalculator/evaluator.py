import logging
from functools import lru_cache

from pandas import DataFrame
from numpy import ndarray
from xlfunctions import xl, xlerrors

from . import xltypes


class Evaluator:
    """Traverses and evaluates a given model."""

    def __init__(self, model):
        self.model = model
        self.recursed_cells = set()
        self.cache_count = 0
        self.cells_to_evaluate = {}

    def recurse_evaluate(
            self, cells, data, cells_to_evaluate, recursed_cells,
            recurse_depth=0
    ):
        """Determine the execution dependency tree.

        Sometimes (maybe most of the time) we need to 'chase' the dependency
        tree to eval the correct value of a cell.
        """
        if len(data) == 0:
            return

        # Base case
        if len(data) == 1 and isinstance(data[0], str):
            if data[0] in cells and cells[data[0]].formula is None:
                if recurse_depth in cells_to_evaluate:
                    cells_to_evaluate[recurse_depth].add(data[0])
                else:
                    for level in cells_to_evaluate:
                        if data[0] in cells_to_evaluate[level]:
                            cells_to_evaluate[level].remove(data[0])
                    cells_to_evaluate[recurse_depth] = set(data)
                    recursed_cells.add(data[0])

        # Recursive cases
        elif (
                len(data) == 1 and isinstance(data[0], str)
                and cells[data[0]].formula is not None
        ):
            self.recurse_evaluate(
                cells,
                list(cells[data[0]].formula.associated_cells),
                cells_to_evaluate, recursed_cells,
                recurse_depth=recurse_depth + 1
            )

            if recurse_depth in cells_to_evaluate:
                cells_to_evaluate[recurse_depth].add(data[0])
                recursed_cells.add(data[0])
            else:
                for level in cells_to_evaluate:
                    if data[0] in cells_to_evaluate[level]:
                        cells_to_evaluate[level].remove(data[0])
                cells_to_evaluate[recurse_depth] = set(data)
                recursed_cells.add(data[0])

        elif len(data) == 1 and isinstance(data[0], list):
            self.recurse_evaluate(
                cells, data, cells_to_evaluate, recursed_cells,
                recurse_depth=recurse_depth + 1)

        else:
            mid = len(data) // 2
            first_half = data[:mid]
            second_half = data[mid:]
            self.recurse_evaluate(
                cells, first_half, cells_to_evaluate, recursed_cells,
                recurse_depth=recurse_depth
            )
            self.recurse_evaluate(
                cells, second_half, cells_to_evaluate, recursed_cells,
                recurse_depth=recurse_depth
            )

    @lru_cache(maxsize=None)
    def evaluate(self, cell_address, clear_cache=False):
        """Evaluates the cell's formula."""
        ns = {}
        ns.update(xl.FUNCTIONS)

        defined_names = self.model.defined_names
        cells = self.model.cells
        ranges = self.model.cells

        if clear_cache:
            self.evaluate.cache_clear()
            self.eval_ref.cache_clear()

        try:
            # Although defined names have been resolved in Model.create_node()
            # we need to attempt to resolve defined names as we might have been
            # given one in argument cell_address.
            if cell_address in defined_names:
                name_definition = defined_names[cell_address]
                if isinstance(name_definition, xltypes.XLCell):
                    cell_address = name_definition.address

                elif isinstance(name_definition, xltypes.XLRange):
                    message = (
                        f"I can't resolve {cell_address} to a cell. It's a "
                        f"range and they aren't supported yet."
                    )
                    logging.error(message)
                    raise ValueError(message)

                elif isinstance(name_definition, xltypes.XLFormula):
                    message = (
                        f"I can't resolve {cell_address} to a cell. It's a "
                        f"formula and they aren't supported as a cell "
                        f"reference."
                    )
                    logging.error(message)
                    raise ValueError(message)

            elif cell_address in ranges:
                cell = ranges[cell_address]

            else:
                cell = cells[cell_address]  # noqa

        except (KeyError, ValueError):
            logging.error('Empty cell at {}'.format(cell_address))
            return xlerrors.NullExcelError(f'Cell {cell_address} is empty')

        # No formula, or no evaluation means fixed value but could be a defined
        # name.
        if (
                cells[cell_address].formula is None
                or cells[cell_address].formula.evaluate is False
        ):
            logging.debug(
                f"Cell {cells[cell_address].address} has no formula "
                f"but its value is {cells[cell_address].value}")
            return cells[cell_address].value

        try:
            if cells[cell_address].formula.ast is not None:
                skip = False
                for level in self.cells_to_evaluate:
                    if cell_address in self.cells_to_evaluate[level]:
                        skip = True

                if not skip:
                    self.recurse_evaluate(
                        cells,
                        list(cells[cell_address].formula.associated_cells),
                        self.cells_to_evaluate, self.recursed_cells
                    )
                    cells_to_evaluate_keys = list(
                        self.cells_to_evaluate.keys())
                    sorted(cells_to_evaluate_keys)
                    cells_to_evaluate_keys.reverse()
                    for item in cells_to_evaluate_keys:
                        for recursed_cell_addr in self.cells_to_evaluate[item]:
                            active_cell = cells[recursed_cell_addr]
                            if active_cell.formula is not None:
                                value = active_cell.formula.ast.eval(
                                    self.model, ns, cell_address)
                                cells[recursed_cell_addr].value = value

                value = cells[cell_address].formula.ast.eval(
                    self.model, ns, cell_address)
                if isinstance(value, ndarray):
                    cells[cell_address].value = value \
                        if len(value) != 0 else None
                else:
                    if isinstance(value, DataFrame):
                        cells[cell_address].value = value \
                            if not value.empty else None
                    else:
                        cells[cell_address].value = value \
                            if value != '' else None

            else:
                cells[cell_address].value = 0

            cells[cell_address].need_update = False

        except Exception as err:
            if str(err).startswith("Problem evalling"):
                raise err
            else:
                raise RuntimeError(
                    "Problem evalling: {} for {}, {}".format(
                        err,
                        cells[cell_address].address,
                        cells[cell_address].formula.formula
                    )
                )

        logging.debug(
            "Cell {} has a formula, {} \r\n"
            "which evaluates to {}\r\n".format(
                cells[cell_address].address,
                cells[cell_address].formula.formula,
                cells[cell_address].value)
        )

        return cells[cell_address].value

    def find_associated_cell(self, ref, range):
        """This function retrieves the cell associated to ref in a Range

        For instance, in the range [A1, B1, C1], the cell associated to B2 is
        B1 This is useful to mimic the way Excel works.
        """
        if ref is None:
            return None

        row, col = ref

        # if a Range is empty, it means normally that all its cells are
        # empty
        if (range.length) == 0:
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

    def find_associated_value(self, ref, item):
        """retrieves the value and not the Cell."""

        row, col = ref

        if isinstance(item, Evaluator):
            try:
                # If a Range is empty, it means normally that all its cells
                # are empty
                if (item.length) == 0:
                    item_value = 0

                elif item.type == "vertical":
                    if item.__cellmap is not None:
                        item_value = item[(row, item.origin[1])].value \
                            if (row, item.origin[1]) in item else None

                    else:
                        item_value = item[(row, item.origin[1])] \
                            if (row, item.origin[1]) in item else None

                elif item.type == "horizontal":

                    if item.__cellmap is not None:
                        item_value = item[(item.origin[0], col)].value \
                            if (item.origin[0], col) in item else None
                    else:
                        item_value = item[(item.origin[0], col)] \
                            if (item.origin[0], col) in item else None

                else:
                    return xlerrors.ValueExcelError(
                        f'cannot use find_associated_value on {item.type}')

            except xlerrors.ExcelError as err:
                raise RuntimeError(
                    f'First argument of Range operation is not '
                    f'valid: {err.value}'
                )

        elif item is None:
            item_value = 0

        else:
            item_value = item

        return item_value

    def set_cell_value(self, address, value, clear_cache=True):
        """Sets the value of a cell in the model."""
        self.model.set_cell_value(address, value)
        if clear_cache:
            self.evaluate.cache_clear()

    def get_cell_value(self, address):
        """Gets the value of a cell in the model."""
        return self.model.get_cell_value(address)
