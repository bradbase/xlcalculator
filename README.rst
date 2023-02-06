================
Excel Calculator
================


.. image:: https://travis-ci.org/bradbase/xlcalculator.png?branch=master
   :target: https://travis-ci.org/bradbase/xlcalculator

.. image:: https://coveralls.io/repos/github/bradbase/xlcalculator/badge.svg?branch=master
   :target: https://coveralls.io/github/bradbase/xlcalculator?branch=master

.. image:: https://img.shields.io/pypi/v/xlcalculator.svg
    :target: https://pypi.python.org/pypi/xlcalculator

.. image:: https://img.shields.io/pypi/pyversions/xlcalculator.svg
    :target: https://pypi.python.org/pypi/xlcalculator/

.. image:: https://img.shields.io/pypi/status/xlcalculator.svg
    :target: https://pypi.org/project/xlcalculator/

xlcalculator is a Python library that reads MS Excel files and, to the extent
of supported functions, can translate the Excel functions into Python code and
subsequently evaluate the generated Python code. Essentially doing the Excel
calculations without the need for Excel.

xlcalculator is a modernization of the
`koala2 <https://github.com/vallettea/koala>`_ library.

``xlcalculator`` currently supports:

* Loading an Excel file into a Python compatible state
* Saving Python compatible state
* Loading Python compatible state
* Ignore worksheets
* Extracting sub-portions of a model. "focussing" on provided cell addresses
  or defined names
* Evaluating

    * Individual cells
    * Defined Names (a "named cell" or range)
    * Ranges
    * Shared formulas `not an Array Formula <https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula>`_

      * Operands (+, -, /, \*, ==, <>, <=, >=)
      * on cells only

    * Set cell value
    * Get cell value
    * `Parsing a dict into the Model object <https://stackoverflow.com/questions/31260686/excel-formula-evaluation-in-pandas/61586912#61586912>`_

        * Code is in examples\\third_party_datastructure

    * Functions are at the bottom of this README

        * LN
            - Python Math.log() differs from Excel LN. Currently returning
              Math.log()

        * VLOOKUP
          - Exact match only

        * YEARFRAC
          - Basis 1, Actual/actual, is only within 3 decimal places

Not currently supported:

  * Array Formulas or CSE Formulas (not a shared formula): https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula or https://support.office.com/en-us/article/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7#ID0EAAEAAA=Office_2013_-_Office_2019)

      * Functions required to complete testing as per Microsoft Office Help
        website for SQRT and LN
      * EXP
      * DB

Run tests
---------

Setup your environment::

  virtualenv -p 3.10 ve
  ve/bin/pip install -e .[test]

From the root xlcalculator directory::

  ve/bin/py.test -rw -s --tb=native

Or simply use ``tox``::

  tox


Run Example
-----------

From the examples/common_use_case directory::

  python use_case_01.py



Adding/Registering Excel Functions
----------------------------------

Excel function support can be easily added.

Fundamental function support is found in the xlfunctions directory. The
functions are thematically organised in modules.

Excel functions can be added by any code using the
``xlfunctions.xl.register()`` decorator. Here is a simple example:

.. code-block:: Python

  from xlcalculator.xlfunctions import xl

  @xl.register()
  @xl.validate_args
  def ADDONE(num: xl.Number):
      return num + 1

The `@xl.validate_args` decorator will ensure that the annotated arguments are
converted and validated. For example, even if you pass in a string, it is
converted to a number (in typical Excel fashion):

.. code-block:: Python

  >>> ADDONE(1):
  2
  >>> ADDONE('1'):
  2

If you would like to contribute functions, please create a pull request. All
new functions should be accompanied by sufficient tests to cover the
functionality. Tests need to be written for both the Python implementation of
the function (tests/xlfunctions) and a comparison with Excel
(tests/xlfunctions_vs_excel).



Excel number precision
----------------------

Excel number precision is a complex discussion.

It has been discussed in a `Wikipedia
page <https://en.wikipedia.org/wiki/Numeric_precision_in_Microsoft_Excel>`_.

The fundamentals come down to floating point numbers and a contention between
how they are represented in memory Vs how they are stored on disk Vs how they
are presented on screen. A `Microsoft
article <https://www.microsoft.com/en-us/microsoft-365/blog/2008/04/10/understanding-floating-point-precision-aka-why-does-excel-give-me-seemingly-wrong-answers/>`_
explains the contention.

This project is attempting to take care while reading numbers from the Excel
file to try and remove a variety of representation errors.

Further work will be required to keep numbers in-line with Excel throughout
different transformations.

From what I can determine this requires a low-level implementation of a
numeric datatype (C or C++, Cython??) to replicate its behaviour. Python
built-in numeric types don't replicate behaviours appropriately.


Unit testing Excel formulas directly from the workbook.
-------------------------------------------------------

If you are interested in unit testing formulas in your workbook, you can use
`FlyingKoala <https://github.com/bradbase/flyingkoala>`_. An example on how can
be found
`here <https://github.com/bradbase/flyingkoala/tree/master/flyingkoala/unit_testing_formulas>`_.


TODO
----

- Do not treat ranges as a granular AST node it instead as an operation ":" of
  two cell references to create the range. That will make implementing
  features like ``A1:OFFSET(...)`` easy to implement.

- Support for alternative range evaluation: by ref (pointer), by expr (lazy
  eval) and current eval mode.

    * Pointers would allow easy implementations of functions like OFFSET().

    * Lazy evals will allow efficient implementation of IF() since execution
      of true and false expressions can be delayed until it is decided which
      expression is needed.

- Implement array functions. It is really not that hard once a proper
  RangeData class has been implemented on which one can easily act with scalar
  functions.

- Improve testing

- Refactor model and evaluator to use pass-by-object-reference for values of
  cells which then get "used"/referenced by ranges, defined names and formulas

- Handle multi-file addresses

- Improve integration with pyopenxl for reading and writing files `example of
  problem space <https://stackoverflow.com/questions/40248564/pre-calculate-excel-formulas-when-exporting-data-with-python>`_



Supported Functions
-------------------


Compatibility
-------------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | FLOOR           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+


Date and Time
-------------

  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | DATE            |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | DATEDIF         |      *       |       |          |       |
  +-----------------+--------------+-------+----------+-------+
  | DATEVALUE       |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DAY             |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DAYS            |      *       |       |          |       |
  +-----------------+--------------+-------+----------+-------+
  | EDATE           |      *       |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | EOMONTH         |      *       |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | HOUR            |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ISOWEEKNUM      |      *       |       |          |       |
  +-----------------+--------------+-------+----------+-------+
  | MINUTE          |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | MONTH           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | NOW             |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SECOND          |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TIME            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TIMEVALUE       |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TODAY           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | WEEKDAY         |      *       |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | YEAR            |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | YEARFRAC        |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+


Engineering
-----------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | BIN2DEC         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | BIN2HEX         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | BIN2OCT         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DEC2BIN         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DEC2HEX         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DEC2OCT         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | HEX2BIN         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | HEX2DEC         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | HEX2OCT         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | OCT2BIN         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | OCT2DEC         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | OCT2HEX         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+


Financial
---------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | IRR             |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | NPV             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | PMT             |      *       |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | PV              |      *       |       |          |       |
  +-----------------+--------------+-------+----------+-------+
  | SLN             |      *       |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | VDB             |      *       |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | XIRR            |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | XNPV            |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+


Information
-----------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | ISBLANK         |      *       |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | ISERR           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ISERROR         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ISEVEN          |      *       |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | ISNA            |      *       |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | ISNUMBER        |      *       |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | ISODD           |      *       |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | ISTEXT          |      *       |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | NA              |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+


Logical
-------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | AND             |      *       |   *   |     *    |       |
  +-----------------+--------------+-------+----------+-------+
  | FALSE           |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | IF              |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | IFERROR         |              |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | IFS             |              |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | NOT             |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | OR              |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SWITCH          |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TRUE            |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | XOR             |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+


Lookup and reference
--------------------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | CHOOSE          |      *       |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | COLUMN          |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | COLUMNS         |              |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | HLOOKUP         |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | INDEX           |              |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | INDIRECT        |              |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | LOOKUP          |              |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | MATCH           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | OFFSET          |              |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | ROW             |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ROWS            |              |       |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | VLOOKUP         |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+


Math and Trigonometry
---------------------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | ABS             |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ACOS            |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ACOSH           |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ACOT            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ACOTH           |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ARABIC          |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ASIN            |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ASINH           |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ATAN            |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ATAN2           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ATANH           |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | CEILING         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | CEILING.MATH    |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | CEILING.PRECISE |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | COS             |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | COSH            |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | COT             |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | COTH            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | CSC             |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | CSCH            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DECIMAL         |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | DEGREES         |     *        |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | EVEN            |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | EXP             |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | FACT            |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | FACTDOUBLE      |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | FLOOR.MATH      |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | FLOOR.PRECISE   |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | GCD             |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | INT             |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ISO.CEILING     |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LCM             |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LN              |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LOG             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | LOG10           |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | MOD             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | MROUND          |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ODD             |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | PI              |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | POWER           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | RADIANS         |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | RAND            |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | RANDBETWEEN     |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | ROMAN           |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ROUND           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | ROUNDDOWN       |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | ROUNDUP         |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | SEC             |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SECH            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SIGN            |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SIN             |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SINH            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SQRT            |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | SQRTPI          |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | SUM             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | SUMIF           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | SUMIFS          |      *       |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | SUMPRODUCT      |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | TAN             |      *       |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TANH            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TRUNC           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+


Statistical
-----------


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | AVERAGE         |      *       |   *   |     *    |   *   |
  +-----------------+--------------+-------+----------+-------+
  | AVERAGEA        |              |       |     *    |       |
  +-----------------+--------------+-------+----------+-------+
  | AVERAGEIF       |              |   *   |     *    |       |
  +-----------------+--------------+-------+----------+-------+
  | AVERAGEIFS      |              |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | COUNT           |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | COUNTA          |      *       |       |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | COUNTBLANK      |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | COUNTIF         |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | COUNTIFS        |      *       |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | LARGE           |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LINEST          |              |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
  | MAX             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | MAXA            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | MAXIFS          |              |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | MIN             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | MINA            |              |       |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | MINIFS          |              |   *   |          |       |
  +-----------------+--------------+-------+----------+-------+
  | SMALL           |              |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+


Text
----


  +-----------------+--------------+-------+----------+-------+
  | Function        | xlcalculator | PyCel | formulas | Koala |
  +-----------------+--------------+-------+----------+-------+
  | CONCAT          |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | CONCATENATE     |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | EXACT           |      *       |       |          |       |
  +-----------------+--------------+-------+----------+-------+
  | FIND            |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LEFT            |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LEN             |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | LOWER           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | MID             |      *       |   *   |    *     |   *   |
  +-----------------+--------------+-------+----------+-------+
  | REPLACE         |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | RIGHT           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | TRIM            |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | UPPER           |      *       |   *   |    *     |       |
  +-----------------+--------------+-------+----------+-------+
  | VALUE           |              |   *   |          |   *   |
  +-----------------+--------------+-------+----------+-------+
