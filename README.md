Links
=====

OfficeGenerator doxygen documentation:

  * http://turulomio.users.sourceforge.net/doxygen/officegenerator/
 
Pypi project page

  * https://pypi.org/project/officegenerator/

Dependencies
============

* https://www.python.org/, as the main programming language.
* https://pypi.org/project/odfpy/, to generate LibreOffice documents.
* https://pypi.org/project/openpyxl/, to generate MS Office XLSX/XLSM  documents.
* http://xmlsoft.org/, to execute xmllint in officegenerator_odf2xml

Usage
=====
You can view officegenerator/demo.py to see an example of code: https://raw.githubusercontent.com/Turulomio/officegenerator/master/officegenerator/demo.py

Known issues
============
  * Search and replace in odf files doesn't work with odfpy-1.4.1, setup forces to use odf.py-1.3.6

Changelog
=========
1.24.0
------
  * Added skip_up and skip_down parameters in ODS_Read and XLSX_Read 'values()' methods to skip rows as necessary.
  * Added method 'values_by_range' in XLSX_Read and ODS_Read.
  * XLSX_Write and XLSX_Read have a new constructor paramter 'data_only' to show last save values of each cell instead of formulas.

1.23.0
------
  * Huge performance improvement reading ODS files.
  * Added skip_down to getColumnValues method and skip_right to getRowValues method.
  * Fixed bug getting string cells in ods.

1.22.0
------
  * Added Model and XLSX_Read to module visibility.
  * Added Model totals to xlsx and ods files.
  * [XLSX] overwrite_formula class type parameter is now a string to avoid innecesary imports.
  * Improved formula styles and types (Money, Currency, bool).
  * Added Model_Auto method for fast sheets.
  * You can now generate a file from a Model method.
  * Faster demo generation.
  * Added standard sheets demos.

1.21.0
------
  * Refactorized OpenPyXL to XLSX_Write.
  * Added XLSX_Read class.
  * Improved demo readonly methods.

1.20.0
------
  * Method created to get column and row values

1.19.0
------
  * Added method to iterate ranges and ods sheets.

1.18.0
------
  * Added suport for time objects
  * Added demo file to search and replace in ODT
  * Demo works on windows again

1.17.0
------
  * Added methods to remove rows and columns in Model class.
  * Improved odt_table for standard_sheets.

1.16.0
------
  * Added boolean number and styles.
  * Added output dir parameter to 'convert_to_pdf' method.
  * Added vertical header width feature

1.15.1
------
  * Added forgotten object files.

1.15.0
------
  * Percentage and Currency objects are now from reusingcode project.
  * Added Xulpymoney Money class support as a string.
  * Model updated to work fine with basic ODS, ODT and XLSX data sheets.  

1.14.0
------
  * Improving search encapsulating types with Strings.

1.13.0
------
  * [#18] Added default topleftcell in freezeAndSelect methods.
  * [#19] Added addRowCopy and addColumnCopy functions.

1.12.0
------
  * Fixing problems with freeze and select
  * Added formulas with different styles
  * Added fist version of Model class and standard sheets
  * Replaced setCursorPosition and setSplitPosition by freezeAndSelect
  * Added officegenerartor_xlsx2xml to help debuging

1.11.0
------
  * Fixing problems with freeze and select
 
1.10.0
------
  * Datetime and date cells are now aligned right by default
  * Added creationdate, description, and keywords to metadata in ODF class

1.9.0
-----
  * Added overwrite_formula method. Now we can define formula number_format

1.8.0
-----
  * Solved bug saving with template
  * Unified method to freeze and set position

1.7.0
-----
  * Now integers in xlsx show thousands separator.
  * Cell number format shows currency and decimals in float numbers, correctly

1.6.1
-----
  * Project moved from Sourceforge to Github

1.6.0
-----
  * [#29] Added function max_rows max_columns

1.5.0
-----
  * Solved bug converting to pdf files with white spaces

1.4.0
-----
  * [#28] Removed deprecation warning when removing OpenPyXL Sheets. Added test to check it.
  * [#19] Added illustration method to add a paragraph with a list of images with the same size.

1.3.0
-----
  * Style None now preserve template style
  * Used libreoffice number formats in libxlsxgenerator

1.2.1
-----
  * Now add in odssheet can guess the style if a color is passed
  * Solved bug adding lists to sheets

1.2.0
-----
  * Added ODT_Standard and ODT_Manual_Styles to package visibility

1.1.0
-----
  * ODT tables now work again
  * Images and tables can be named now in ODT documents
  * Added subtitle, bold, underlined and illustrator styles in ODT documents
  * Added cursor to ODT to add Elements wherever you want
  * Added convert_to_pdf in ODT

1.0.0
-----
  * [#7] Solved bug with charmap in Windows
  * [#10] Dependencies are installed when using pip
  * [#11] Now you can add diferent currencies
  * [#13] Now cells are vertical aligned
  * [#14] Added a normal style to predefined colors
  * [#15] officegenerator_demo --remove works now in Windows
  * [#16] Moved missing functions with letter, number args to coord 
  * python setup.py uninstall works now in Windows

0.13.0
------
  * Coord and Range addRow/Column functions now updates the object and doesn't create a new one

0.12.0
------
  * Improving Range and Coord limit situations

0.11.0
------
  * Added package tests
  * Added append/prepend rows/columns to Range
  * Added guess_ods_style function

0.10.0
------
  * Replaced letter, number parameters by Coord and Range
  * Added compatibilty classes OpenPyXL2010 and ODS_Write_Without_Styles
  * Colors and data styles work in XLSX and ODS

0.9.0
-----
  * Solved problem with freeze and setselectedcell

0.8.0
-----
  * __version__ is now in commons
  * Improved spanish translation
  * Solved problem with merged cells
  * Added overwrite_and_merge to XLSX generator

0.7.0
-----
  * Replaced predefined styles by predefined colors
  * Added officegenerator.xlsx to demo

0.6.0
-----
  * Addapted code to openpyxl-2.4.1

0.5.0
-----
  * [#1] Added dependencies in setup.py
  * [#5] officegenerartor_odf2xml require --file parameter
  * [#6] Show alerts in Windows if can't be executed
  * Added internationalization infrastructure

0.4.0
-----
  * Added pkg_resources support
  * Moved images directory to package

0.3.1
-----
  * Solved bug with image path in demo.py

0.3.0
-----
  * Now officegenerator_demo can delete example files with --remove parameter
  * Added officegenerator_odf2xml to convert odf files to indented xml

0.2.0
-----
  * Added officegenerator_demo to view basic examples and code

0.1.0
-----
  * Basic funcionality
