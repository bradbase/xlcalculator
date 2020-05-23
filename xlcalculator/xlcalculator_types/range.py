"""Representation of a Microsoft Excel range
"""
from dataclasses import dataclass, field

from .xltype import XLType
from . import utils


@dataclass
class XLRange(XLType):
    """"""

    address_str: str = field(compare=False, hash=False, repr=True)
    name: str = field(default=None, compare=False, hash=True, repr=True)

    cells: list = field(init=False, compare=True, hash=False, repr=False)
    sheet: str = field(init=False, default="Sheet1", repr=False)
    value: list = field(default=None, repr=True)

    def __post_init__(self):
        if self.name is None:
            self.name = self.address_str
        self.sheet, self.cells = utils.resolve_ranges(self.address_str)

    @property
    def address(self):
        return self.cells
