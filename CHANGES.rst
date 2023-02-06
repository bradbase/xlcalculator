=======
CHANGES
=======

0.5.0 (2023-02-06)
------------------

- Added support for Python 3.10, dropped 3.8.

- Upgraded requirements.txt to latest versions.

  * `yearfrac==0.4.4` was incompatible with latest setuptools.

  * `openpyxl` had API changes that were addressed and tests fixed.


0.4.2 (2021-05-17)
------------------

- Make sure that decimal rounding is only set in context and not system wide.

0.4.1 (2021-05-14)
------------------

- Fixed cross-sheet references.


0.4.0 (2021-05-13)
------------------

- Pass ``ignore_hidden`` from ``read_and_parse_archive()`` to
  ``parse_archive()``

- Add Excel tests for ``IF()``.

- Add ``NOT()`` function.

- Implemented ``BIN2OCT()``, ``BIN2DEC()``, ``BIN2HEX()``, ``OCT2BIN()``,
  ``OCT2DEC()``, ``OCT2HEX()``, ``DEC2BIN()``, ``DEC2OCT()``, ``DEC2HEX()``,
  ``HEX2BIN()``, ``HEX2OCT()``, ``HEX2DEC()``.

- Drop Python 3.7 support.


0.3.0 (2021-05-13)
------------------

- Add support for cross-sheet references.

- Make ``*IF()`` functions case insensitive to properly adhere to Excel specs.

- Support for Python 3.9.


0.2.13 (2020-12-02)
-------------------

- Add functions: ``FALSE()``, ``TRUE()``, ``ATAN2()``, ``ACOS()``,
  ``DEGREES()``, ``ARCCOSH()``, ``ASIN()``, ``ASINH()``, ``ATAN()``,
  ``CEILING()``, ``COS()``, ``RADIANS()``, ``COSH()``, ``EXP()``, ``EVEN()``,
  ``FACT()``, ``FACTDOUBLE()``, ``INT()``, ``LOG()``, ``LOG10()``. ``RAND()``,
  ``RANDBETWRRN()``, ``SIGN()``, ``SIN()``, ``SQRTPI()``, ``TAN()``


0.2.12 (2020-11-28)
-------------------

- Add functions: ``PV()``, ``XIRR()``, ``ISEVEN()``, ``ISODD()``,
  ``ISNUMBER()``, ``ISERROR()``, ``FLOOR()``, ``ISERR()``
- Bugfix unary operator needed to be right associated to handle cases of
  double use eg; double-negative.. --4 == 4


0.2.11 (2020-11-16)
-------------------

- Add functions: ``DAY()``, ``YEAR()``, ``MONTH()``, ``NOW()``, ``WEEKDAY()``
  ``EDATE()``, ``EOMONTH()``, ``DAYS()``, ``ISOWEEKNUM()``, ``DATEDIF()``
  ``FIND()``, ``LEFT()``, ``LEN()``, ``LOWER()``, ``REPLACE()``, ``TRIM()``
  ``UPPER()``, ``EXACT()``


0.2.10 (2020-10-30)
-------------------

- Support CONCATENATE
- Update setup.py classifiers, licence and keywords


0.2.9 (2020-09-26)
------------------

- Bugfix ModelCompiler.read_and_parse_dict() where a dict being parsed into a
  Model through ModelCompiler was triggering AttributeError on calling
  xlcalculator.xlfunctions.xl. It's a leftover from moving xlfunctions into
  xlcalculator. There has been a test included.


0.2.8 (2020-09-22)
------------------

- Fix implementation of ``ISNA()`` and ``NA()``.

- Impement ``MATCH()``.


0.2.7 (2020-09-22)
------------------

- Add functions: ``ISBLANK()``, ``ISNA()``, ``ISTEXT()``, ``NA()``


0.2.6 (2020-09-21)
------------------

- Add ``COUNTIIF()`` and ``COUNTIFS()`` function support.


0.2.5 (2020-09-21)
------------------

- Add ``SUMIFS()`` support.


0.2.4 (2020-09-09)
------------------

- Updated README with supported functions.

- Fix bug in ModelCompiler extract method where a defined name cell was being
  overwritten with the cell from one of the terms contained within the formula.
  Added a test for this.

- Move version of yearfrac to 0.4.4. That project has removed a dependency
  on the package six.


0.2.3 (2020-08-18)
------------------

- In-boarded xlfunctions.

- Bugfix COUNTA.

  * Now supports 256 arguments.

- Updated README. Includes words on xlfunction.

- Changed licence from GPL-3 style to MIT Style.


0.2.2 (2020-05-28)
------------------

- Make dependency resolution part of the execution.

  * AST eval'ing takes care of depedency resolution.

  * Provide cycle detection with reporting.

  * Implemented a specific evaluation context. That makes cache control,
    namespace customization and data encapsulation much easier.

- Add more tokenizer tests to increase coverage.


0.2.1 (2020-05-28)
------------------

- Use a less intrusive way to patch ``openpyxl``. Instead of permanently
  patching the reader to support cached formula values, ``mock`` is used to
  only patch the reader while reading the workbook.

  This way the patches do not interfere with other packages not expecting
  these new classes.


0.2.0 (2020-05-28)
------------------

- Support for delayed node evaluation by wrapping them into expressions. The
  function will eval the expression when needed.

- Support for native Excel data types.

- Enable and update Excel file based function tests that are now working
  properly.

- Flake8 source code.


0.1.0 (2020-05-25)
------------------

- Refactored ``xlcalculator`` types to be more compact.

- Reimplemented evaluation engine to not generate Python code anymore, but
  build a proper AST from the AST nodes. Each AST node supports an `eval()`
  function that knows how to compute a result.

  This removes a lot of complexities around trying to determine the evaluation
  context at code creation time and encoding the context as part of the
  generated code.

- Removal of all special function handling.

- Use of new `xlfunctions` implementation.

- Use Openpyxl to load the Excel files. This provides shared formula support
  for free.


0.0.1b (2020-05-03)
-------------------

- Initial release.
