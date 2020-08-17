import datetime
import dateutil
import numpy
import pandas
from typing import Optional, Union, NewType

from . import utils, xlerrors

NATIVE_TO_XLTYPE = {}


def register(cls):
    for native_type in cls.native_types:
        NATIVE_TO_XLTYPE[native_type] = cls
    return cls


class ExcelType:

    __slots__ = ('value')

    native_types = ()

    sort_precedence = 0

    def __new__(cls, value):
        inst = super().__new__(cls)
        assert isinstance(value, cls.native_types), value
        inst.value = value
        return inst

    @classmethod
    def cast(cls, value):
        if isinstance(value, cls):
            return value
        if not isinstance(value, ExcelType):
            if type(value) not in NATIVE_TO_XLTYPE:
                raise xlerrors.ValueExcelError(
                    f'Unknown object type: {type(value)} ({value})')
            value = NATIVE_TO_XLTYPE[type(value)](value)
        return getattr(value, f'__{cls.__name__}__')()

    @classmethod
    def is_type(cls, value):
        return isinstance(value, (cls,) + cls.native_types)

    @classmethod
    def cast_from_native(cls, value):
        if isinstance(value, xlerrors.ExcelError):
            return value
        if isinstance(value, tuple(NATIVE_TO_XLTYPE.values())):
            return value
        return NATIVE_TO_XLTYPE[type(value)](value)

    def _sort_key(self, other):
        return (self.sort_precedence, self.value)

    def __add__(self, other):
        return Number(Number.cast(self).value + Number.cast(other).value)

    def __sub__(self, other):
        return Number(Number.cast(self).value - Number.cast(other).value)

    def __mul__(self, other):
        return Number(Number.cast(self).value * Number.cast(other).value)

    def __truediv__(self, other):
        ovalue = float(Number.cast(other))
        if ovalue == 0:
            raise xlerrors.DivZeroExcelError()
        return Number(float(Number.cast(self)) / ovalue)

    def __pow__(self, other):
        return Number(Number.cast(self).value ** Number.cast(other).value)

    def __and__(self, other):
        # Highjacking bitwise "and" to implement logical "and"
        return Boolean(bool(self) and bool(other))

    def __or__(self, other):
        # Highjacking bitwise "or" to implement logical "or"
        return Boolean(bool(self) or bool(other))

    __radd__ = __add__
    __rsub__ = __sub__
    __rmul__ = __mul__
    __rtruediv__ = __truediv__
    __rpow__ = __pow__
    __rand__ = __and__
    __ror__ = __or__

    def __lt__(self, other):
        other = ExcelType.cast_from_native(other)
        return Boolean(self._sort_key(other) < other._sort_key(self))

    def __le__(self, other):
        other = ExcelType.cast_from_native(other)
        return Boolean(self._sort_key(other) <= other._sort_key(self))

    def __eq__(self, other):
        other = ExcelType.cast_from_native(other)
        return Boolean(self._sort_key(other) == other._sort_key(self))

    def __ne__(self, other):
        other = ExcelType.cast_from_native(other)
        return Boolean(self._sort_key(other) != other._sort_key(self))

    def __gt__(self, other):
        other = ExcelType.cast_from_native(other)
        return Boolean(self._sort_key(other) > other._sort_key(self))

    def __ge__(self, other):
        other = ExcelType.cast_from_native(other)
        return Boolean(self._sort_key(other) >= other._sort_key(self))

    def __int__(self):
        try:
            return int(float(self.value))
        except ValueError:
            raise xlerrors.ValueExcelError(
                f'Could not convert {repr(self.value)} to int.')

    def __float__(self):
        try:
            return float(self.value)
        except (TypeError, ValueError):
            raise xlerrors.ValueExcelError(
                f'Could not convert {repr(self.value)} to float.')

    def __str__(self):
        return str(self.value)

    def __bool__(self):
        return bool(self.value)

    def __number__(self):
        return float(self.value)

    def __datetime__(self):
        raise NotImplementedError

    def __Number__(self):
        return Number(self.__number__())

    def __Text__(self):
        return Text(self.__str__())

    def __Boolean__(self):
        return Boolean(self.__bool__())

    def __DateTime__(self):
        return DateTime(self.__datetime__())

    def __Blank__(self):
        return self.__class__(None)

    def __hash__(self):
        return hash(self.value)

    def __repr__(self):
        return f'<{self.__class__.__name__} {repr(self.value)}>'


@register
class Number(ExcelType):

    native_types = (int, float, numpy.int64, numpy.float64)

    blank_value = 0

    @property
    def is_whole(self):
        return isinstance(self.value, int)

    @property
    def is_decimal(self):
        return isinstance(self.value, float)

    def __mod__(self, other):
        return Number(self.value % Number.cast(other).value)

    __rmod__ = __mod__

    def __neg__(self):
        return Number(self.value.__neg__())

    def __pos__(self):
        return Number(self.value.__pos__())

    def __invert__(self):
        return Number(-self.value)

    def __abs__(self):
        return Number(self.value.__abs__())

    def __round__(self, ndigits=None):
        return Number(self.value.__round__(ndigits))

    def __trunc__(self):
        return Number(self.value.__trunc__())

    def __number__(self):
        return self.value

    def __datetime__(self):
        return utils.number_to_datetime(self.value)

    def __Blank__(self):
        return self.__class__(0)


@register
class Text(ExcelType):

    boolean_texts = ['false', 'true']

    native_types = (str,)
    sort_precedence = 1

    def __number__(self):
        try:
            return int(self.value)
        except ValueError:
            pass
        try:
            return float(self.value)
        except ValueError:
            pass
        # For arithmetic, boolean text is actually interpreted.
        try:
            return int(self.__bool__(by_content_only=True))
        except xlerrors.ValueExcelError:
            pass
        # Try casting to datetime first, since the string might represent one
        # and that can be concerted.
        try:
            return utils.datetime_to_number(self.__datetime__())
        except xlerrors.ValueExcelError:
            pass
        raise xlerrors.ValueExcelError(
            f'Could not convert {repr(self.value)} to float.')

    def __bool__(self, by_content_only=False):
        if self.value.lower() in self.boolean_texts:
            return (self.value.lower() == 'true')
        if by_content_only:
            raise xlerrors.ValueExcelError(
                f'Could not convert {repr(self.value)} to bool.')
        return bool(self.value)

    def __Boolean__(self):
        return Boolean(self.__bool__(by_content_only=True))

    def __datetime__(self):
        try:
            return utils.number_to_datetime(float(self.value))
        except (ValueError, OverflowError):
            pass
        try:
            return dateutil.parser.parse(self.value)
        except (ValueError, OverflowError):
            pass
        raise xlerrors.ValueExcelError(
            f'Could not cast {repr(self.value)} (of type {type(self.value)} '
            f'to date/time.')

    def __Blank__(self):
        return self.__class__('')


@register
class Boolean(ExcelType):

    native_types = (bool,)
    sort_precedence = 2
    datetime_true = datetime.datetime(1999, 12, 31)
    datetime_false = datetime.datetime(1999, 12, 30)

    def _sort_key(self, other):
        return (self.sort_precedence, int(self.value))

    def __number__(self):
        return int(self.value)

    def __Boolean__(self):
        return self

    def __datetime__(self):
        return self.datetime_true if self.value else self.datetime_false

    def __Blank__(self):
        return self.__class__(False)


@register
class DateTime(ExcelType):

    native_types = (datetime.datetime, numpy.datetime64)

    def _sort_key(self, other):
        return self.__Number__()._sort_key(other)

    def __int__(self):
        return int(float(self))

    def __float__(self):
        return utils.datetime_to_number(self.value)

    __number__ = __float__

    def __DateTime__(self):
        return self

    def __Blank__(self):
        return None

    def __repr__(self):
        return f'<{self.__class__.__name__} {self.value.isoformat()}>'


@register
class Blank(ExcelType):

    native_types = (type(None),)

    def __new__(cls, value=None):
        return super().__new__(cls, None)

    @classmethod
    def is_blank(cls, value):
        return isinstance(value, (cls,) + cls.native_types) or value == ''

    def _sort_key(self, other):
        return other.__Blank__()._sort_key(self)

    def __and__(self, other):
        if isinstance(other, self.native_types + (Blank,)):
            raise xlerrors.ValueExcelError('Cannot AND two blank values.')
        return Boolean.cast(other)

    def __or__(self, other):
        if isinstance(other, self.native_types + (Blank,)):
            raise xlerrors.ValueExcelError('Cannot OR two blank values.')
        return Boolean.cast(other)

    def __number__(self):
        return 0.0

    def __int__(self):
        return 0

    def __float__(self):
        return 0.0

    def __bool__(self):
        return False

    def __str__(self):
        return ''

    def __repr__(self):
        return '<BLANK>'


BLANK = Blank()


def _safe_cast(func, empty_value=None):

    def safe_cast(value):
        try:
            return func(value)
        except xlerrors.ExcelError:
            return empty_value

    return safe_cast


def _convert_nested_list(value):
    if not isinstance(value, (list, tuple)):
        try:
            return ExcelType.cast_from_native(value)
        except Exception:
            # There are cases where internal pandas data structures are passed
            # around.
            return value
    return [_convert_nested_list(item) for item in value]


@register
class Array(pandas.DataFrame):

    native_types = (list, tuple)

    def __init__(self, data, *args, **kw):
        try:
            super().__init__(_convert_nested_list(data), *args, **kw)
        except ValueError:
            raise xlerrors.ValueExcelError(f'Invalid array argument: {data}')

    @property
    def _constructor(self):
        return Array

    @property
    def flat(self):
        return list(self.values.flat)

    @classmethod
    def cast(cls, value):
        if isinstance(value, cls):
            return value
        if not isinstance(value, cls.native_types):
            value = [[value]]
        return Array(value)

    def flatten(self, xltype=None, filt=None):
        cast = _safe_cast(Number.cast, None) \
            if xltype is not None else lambda x: x
        return list(filter(filt, [cast(item) for item in self.values.flat]))

    def cast_to_numbers(self):
        return self.applymap(_safe_cast(Number.cast, Number(0.0)))

    def cast_to_booleans(self):
        return self.applymap(_safe_cast(Boolean.cast, Boolean(True)))

    def cast_to_texts(self):
        return self.applymap(_safe_cast(Text.cast, Text('')))


class Expr:
    """Expression

    Represents an expression that has yet to be evaluated. This class is
    agnostic to the internals of the implementation details. In other words,
    it does not expect particular objects. all it needs is a callable and its
    arguments.

    The constructor also accepts any arbitrary keywords that will be set on
    the expression to allow for debugging and discovery. Such info can include
    AST nodes, return type, cell reference, etc. (None of th einfo fields are
    required or will be considered by this library.)
    """

    def __init__(self, callable, args=(), kwargs={}, **info):
        self.callable = callable
        self.args = args
        self.kwargs = kwargs.copy()
        for name, value in info.items():
            setattr(self, name, value)

    @classmethod
    def cast(cls, value):
        if isinstance(value, cls):
            return value
        return ValueExpr(value)

    def __call__(self):
        return self.callable(*self.args, *self.kwargs)


def ValueExpr(value):
    return Expr(lambda: value, value=value)


# Python compatible types

# We can really anything into Excel arguments, so let's allow all those
# types.
_Anything = Optional[Union[
    int, float, bool, str, datetime.datetime,
    Number, Text, Boolean, DateTime, Blank, xlerrors.ExcelError
]]

# However, we are going to define specific types nonetheless, so they can be
# used to cast values to proper internal types.
XlNumber = NewType('XlNumber', _Anything)
XlText = NewType('XlText', _Anything)
XlBoolean = NewType('XlBoolean', _Anything)
XlDateTime = NewType('XlDateTime', _Anything)
XlBlank = NewType('XlBlank', _Anything)
XlArray = NewType('XlArray', Union[_Anything, list, Array])
XlExpr = NewType('XlExpr', Union[_Anything, Expr])

# XlAnything declaration for open-ended return type. XlExpr is specifically
# excluded, since that is not a valid return type.
XlAnything = Union[
    XlNumber, XlText, XlBoolean, XlDateTime, XlBlank, XlArray]
