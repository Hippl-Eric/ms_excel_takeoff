# ms_excel_python
Manipulate Microsoft Excel files using Python and [openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html).


Key Features:
- Open, modify, and save Microsoft Excel .xlsx files
- Mimic Excel's insert row functionality by:
    - Translate worksheet formulas 
    - "Drag down" and copy cell values, formulas, and styling
    - Correct sheet row heights
    - Correct merged cells
    - Correct data validation list references

### Table of Contents

- [Background](#background)
- [File Summary](#file-summary)
    - [takeoff.py](#takeoff.py)
    - [helpers.py](#helpers.py)
    - [test_takeoff.py](#test_takeoff.py)
    - [test_helpers.py](#test_helpers.py)
- [What I Learned](#what-i-learned)
- [Contact](#contact)

### Background

While working as a civil engineer estimator, I used a template excel spreadsheet for each of my projects.  At the start of each project I needed to complete the same operations: open template file, update project name, updated cell values, insert rows (sometimes 100's), fix print area, save to new project directory.  Things became very redundant; there had to be a better way.  ms_excel_takeoff was written to automate this entire process.

### File Summary

#### takeoff.py
Main script for modifying a template excel workbook.  Receive user input for project name, number of rows, and other parameters.
#### helpers.py
Helper functions that support takeoff.py.
#### test_takeoff.py
Unit testing for takeoff.py.  Utilizes setUp and tearDown methods to create a test workbook and load a check workbook.  Completes numerous tests utilizing asserts to test values, lists, and dicts.  Custom error messaging is used to pinpoint error locations in files.
#### test_helpers.py
Unit testing for helpers.py.

### What I Learned
- Working with an open source project - Documentation for the project is sparse in areas and required digging into source code to find what I needed.
- Version Control - I started with a MVP and used that in my daily workflow.  As I thought of more features I created branches to separate development and "production" code.  Changes where merged after features were complete and passed testing.
- Unit Testing - To ensure merged features did not break production code, extensive unit testing was used.  Most unit tests were written before the features were built after reviewing project documentation.

### Contact

Eric Hippler, [LinkedIn](https://www.linkedin.com/in/eric-hippler/)
