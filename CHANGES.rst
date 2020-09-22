=======
CHANGES
=======


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
