
# Koala Excel Calculator

In short Koala_xlcalculator is an attempted re-write of the [koala2](https://github.com/vallettea/koala) library with intent to modernize the code, increase maintainability and keep compatibility with external software solutions (self-indulgently one is - [FlyingKoala](https://github.com/bradbase/flyingkoala) ).

Koala2's heritage has been "spreadsheet replacement" in order to get performant calculation. In essence replacing the calculation engine in Excel with one which works faster. And, all credit to the civic hacker crew at [we are ants](https://weareants.fr/#!/koala-the-faster-excel), they absolutely succeeded. In various iterations of the project since it appears as though the koala2 has tried moving toward the idea of something like openpyxl with a "calculate" button.

This is not the direction I've taken.

Koala2 had a few features which enabled me to use koala2 quite differently. I admit these features may not have been core to the wider audience. The features I am specifically interested in are the "advanced" features of setting output cells on load and saving/loading the state of the Python representation of the worksheet. With setting the output cell on load it was possible to load only one formula, and/or the related parent formulas and the input cells involved with the network of formulas. In essence focusing the resulting "loaded Excel" to far less than the entire workbook. And, once in Python, it could be persisted (saved) and re-loaded. This created a mathematical model defined in the language of Excel formulas but executable ("evaluated") and iterable in Python without the need for a person to translate it.

This use case provides the ability to view Excel workbooks as definitions of a calculation (a model or part thereof). Where the model is an abstraction that just so happens to be expressed in the form of an Excel spreadsheet. So, in that sense, koala2 wasn't simply replacing the Excel calculation engine but was more a toolkit which uses Excel as an interface for defining parts or whole mathematical systems. I saw that as very useful.

These features I so covet have broken after some (much needed) code cleanup and I've struggled to re-implement them. With that, I have decided to use some parts of koala2 and implement a library which will have the features I'm interested in. I'm happy if this implementation (or something similar) becomes adopted by the koala2 project (I'd prefer to not split a universe) but, equally, I accept my goals may not overlap enough with that project.

Moving forward, this project, if used in one way, would achieve the purposes of replacing the Excel calculation engine or even be analogous to openpyxl with a "go" button. But would also offer significantly more for those who are inextricably bound to Excel but need more from their math. Which is what FlyingKoala is intended to facilitate.

We have some very basic functionality working. That said, the project is still evolving and has some way to go before it can claim to be feature compatible with koala2.

koala_xlcalculator currently supports:
* Loading an Excel file into a Python compatible state
* Saving Python compatible state
* Loading Python compatible state
* Evaluating
  * Individual cells
  * Defined Names (a "named cell" or range)
  * Ranges
  * Shared formulas (not an Array Formula: https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula)
  * Operands (+, -, /, \*, ==, <>, <=, >=)
    * on cells only
  * Functions
    * AVERAGE
    * CHOOSE
    * CONCAT
    * COUNT
    * COUNTA
    * DATE
    * IRR
    * LN
      - Python Math.log() differs from Excel LN. Currently returning Math.log()
    * MAX
    * MID
    * MIN
    * MOD
    * NPV
    * PMT
    * POWER
    * RIGHT
      - The anthill implementation of Koala was wacky, refers: "hack to deal with naca section numbers" so have introduced an optional ANTHILL compatibility for this method
    * ROUND
    * ROUNDDOWN
    * ROUNDUP
    * SLN
    * SQRT
    * SUM
    * SUMPRODUCT
    * TODAY
    * VLOOKUP
      - Exact match only
    * XNPV
    * YEARFRAC
      - Basis 1, Actual/actual, is only within 3 decimal places
  * Set cell value
  * Get cell value

Not currently supported:
* Array Formulas or CSE Formulas (not a shared formula: https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula or https://support.office.com/en-us/article/guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7#ID0EAAEAAA=Office_2013_-_Office_2019)
* Functions required to complete testing as per Microsoft Office Help website for SQRT and LN
  * ABS
  * EXP
  * DB
* Functions (to be feature complete against Koala2 0.0.31)
  * CONCATENATE
  * COUNTIF
  * COUNTIFS
  * IFERROR
  * INDEX
  * ISBLANK
  * ISNA
  * ISTEXT
  * LINEST
  * LOOKUP
  * MATCH
  * OFFSET
  * VDB

# Run tests
From the root koala_xlcalculator directory
```python
python -m unittest discover -p "*_test.py"
```

# Run Example
From the examples/common_use_case directory
```python
python use_case_01.py
```

# How to add Excel functions
Excel function support can be easily added to koala_xlcalculator.

Do the git things.. fork, clone, branch. checkout the new branch and then;
- Write a class for the function in function_library. Use existing supported function classes as template examples.
- Add the function name and related class to excel_lib.py SUPPORTED_FUNCTIONS dict
- Add the class to function_library\\\_\_init\_\_.py
- Write a test for it in tests\\function_library. Use existing tests as template examples. Ensure you include an associated .xlsx file and implement one or more evaluate test methods. Often a great place for example test ideas is found on the Microsoft Office Excel help page for that function.
- Update the README.md to state that function is supported.
- Put your code, tests and doco forward as a pull request.

# Excel number precision
Excel number precision is a complex discussion.

It has been discussed in a (Wikipedia page)[https://en.wikipedia.org/wiki/Numeric_precision_in_Microsoft_Excel].

The fundamentals come down to floating point numbers and a contention between how they are represented in memory Vs how they are stored on disk Vs how they are presented on screen. A (Microsoft article)[https://www.microsoft.com/en-us/microsoft-365/blog/2008/04/10/understanding-floating-point-precision-aka-why-does-excel-give-me-seemingly-wrong-answers/] explains the contention.

This project is taking care while reading numbers from the Excel file to try and remove a variety of representation errors.

Further work will be required to keep numbers in-line with Excel throughout different transformations.

# TODO
- Fix all functions in the function_library so that they work.
- Set up a travis continuous integration service
- Improve testing
- Refactor model and evaluator to use pass-by-object-reference for values of cells which then get "used"/referenced by ranges, defined names and formulas
- Refactor to ensure the function library only ever gets a non-koala datatype (eg; should only ever get types from pandas, numpy or Python built-in)
