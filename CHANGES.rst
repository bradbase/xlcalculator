=======
CHANGES
=======


0.2.1 (2020-05-28)
------------------

- Use a less intrusive way to patch ``openpyxl``. Instead of permanently
  patching the reader to support cahced formula values, ``mock`` is used to
  only patch the reader while reading the workbook.

  This way the patches do not interfere with other packages not expecting
  these new classes.


0.2.0 (2020-05-28)
------------------

- Support for delayed node evaluation by wrapping them into expressions. The
  fucntion will eval the expression when needed.

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
