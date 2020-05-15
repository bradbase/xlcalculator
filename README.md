
# Excel Calculator

In short xlcalculator is an attempted re-write of the [koala2](https://github.com/vallettea/koala) library with intent to modernize the code, increase maintainability and keep compatibility with external software solutions (self-indulgently one is - [FlyingKoala](https://github.com/bradbase/flyingkoala) ).

Koala2's heritage has been "spreadsheet replacement" in order to get performant calculation. In essence replacing the calculation engine in Excel with one which works faster. And, all credit to the civic hacker crew at [we are ants](https://weareants.fr/#!/koala-the-faster-excel), they absolutely succeeded. In various iterations of the project since it appears as though the koala2 has tried moving toward the idea of something like openpyxl with a "calculate" button.

This is not the direction I've taken.

Koala2 had a few features which enabled me to use koala2 quite differently. I admit these features may not have been core to the wider audience. The features I am specifically interested in are the "advanced" features of setting output cells on load and saving/loading the state of the Python representation of the worksheet. With setting the output cell on load it was possible to load only one formula, and/or the related parent formulas and the input cells involved with the network of formulas. In essence focusing the resulting "loaded Excel" to far less than the entire workbook. And, once in Python, it could be persisted (saved) and re-loaded. This created a mathematical model defined in the language of Excel formulas but executable ("evaluated") and iterable in Python without the need for a person to translate it.

This use case provides the ability to view Excel workbooks as definitions of a calculation (a model or part thereof). Where the model is an abstraction that just so happens to be expressed in the form of an Excel spreadsheet. So, in that sense, koala2 wasn't simply replacing the Excel calculation engine but was more a toolkit which uses Excel as an interface for defining parts or whole mathematical systems. I saw that as very useful.

These features I so covet have broken after some (much needed) code clean-up and I've struggled to re-implement them. With that, I have decided to use some parts of koala2 and implement a library which will have the features I'm interested in. I'm happy if this implementation (or something similar) becomes adopted by the koala2 project (I'd prefer to not split a universe) but, equally, I accept my goals may not overlap enough with that project.

Moving forward, this project, if used in one way, would achieve the purposes of replacing the Excel calculation engine or even be analogous to openpyxl with a "go" button. But would also offer significantly more for those who are inextricably bound to Excel but need more from their math. Which is what FlyingKoala is intended to facilitate.

We have some very basic functionality working. That said, the project is still evolving and has some way to go before it can claim to be feature compatible with koala2.

xlcalculator currently supports:
* Loading an Excel file into a Python compatible state
* Saving Python compatible state
* Loading Python compatible state
* Ignore worksheets
* Extracting sub-portions of a model. "focussing" on provided cell addresses or defined names
* Evaluating
  * Individual cells
  * Defined Names (a "named cell" or range)
  * Ranges
  * Shared formulas [not an Array Formula](https://stackoverflow.com/questions/1256359/what-is-the-difference-between-a-shared-formula-and-an-array-formula)
  * Operands (+, -, /, \*, ==, <>, <=, >=)
    * on cells only
  * Set cell value
  * Get cell value
  * [Parsing a dict into the Model object](https://stackoverflow.com/questions/31260686/excel-formula-evaluation-in-pandas/61586912#61586912)
    * Code is in examples\\third_party_datastructure
  * Functions as implemented in [xlfunctions](https://github.com/bradbase/xlfunctions).
    * ABS
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
From the root xlcalculator directory
```python
python -m unittest discover -p "*_test.py"
```

# Run Example
From the examples/common_use_case directory
```python
python use_case_01.py
```

# How to add Excel functions
Excel function support can be easily added.

Fundamental function support is supplied by [xlfunctions](https://github.com/bradbase/xlfunctions), so to add the "recipe for calculation" please submit a pull request to that project. There are instructions in that project. Please be conscientious with writing tests in that project as they are the tests for _how_ the calculation operates.

It is also best for your submission to have an evaluation test here in xlcalculator so we can ensure that the results of the xlfunction implementation are aligning with what we see in Excel.


# Excel number precision
Excel number precision is a complex discussion. There is further detail on the README at [xlfunctions](https://github.com/bradbase/xlfunctions).


# Unit testing Excel formulas directly from the workbook.

If you are interested in unit testing formulas in your workbook, you can use [FlyingKoala](https://github.com/bradbase/flyingkoala). An example on how can be found [here](https://github.com/bradbase/flyingkoala/tree/master/flyingkoala/unit_testing_formulas).


# TODO
- Set up a travis continuous integration service
- Improve testing
- Refactor model and evaluator to use pass-by-object-reference for values of cells which then get "used"/referenced by ranges, defined names and formulas
- Handle multi-file addresses
- Improve integration with pyopenxl for reading and writing files (Maybe integrating xlcalculator with pyopenxl is actually Koala3?) Example of problem space [here](https://stackoverflow.com/questions/40248564/pre-calculate-excel-formulas-when-exporting-data-with-python)

# BUGS
- Formatted text in a cell (eg; a subscript) breaks the reader.
