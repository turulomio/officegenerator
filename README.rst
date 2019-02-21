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

Changelog
=========
1.7.0
  * Now integers in xlsx show thousands separator.
  * Cell number format shows currency and decimals in float numbers, correctly
1.6.1
  * Project moved from Sourceforge to Github
1.6.0
  * [#29] Added function max_rows max_columns
1.5.0
  * Solved bug converting to pdf files with white spaces
1.4.0
  * [#28] Removed deprecation warning when removing OpenPyXL Sheets. Added test to check it.
  * [#19] Added illustration method to add a paragraph with a list of images with the same size.
1.3.0
  * Style None now preserve template style
  * Used libreoffice number formats in libxlsxgenerator
1.2.1
  * Now add in odssheet can guess the style if a color is passed
  * Solved bug adding lists to sheets
1.2.0
  * Added ODT_Standard and ODT_Manual_Styles to package visibility
1.1.0
  * ODT tables now work again
  * Images and tables can be named now in ODT documents
  * Added subtitle, bold, underlined and illustrator styles in ODT documents
  * Added cursor to ODT to add Elements wherever you want
  * Added convert_to_pdf in ODT
1.0.0
  * [#7] Solved bug with charmap in Windows
  * [#10] Dependencies are installed when using pip
  * [#11] Now you can add diferent currencies
  * [#13] Now cells are vertical aligned
  * [#14] Added a normal style to predefined colors
  * [#15] officegenerator_demo --remove works now in Windows
  * [#16] Moved missing functions with letter, number args to coord 
  * python setup.py uninstall works now in Windows
0.13.0
  * Coord and Range addRow/Column functions now updates the object and doesn't create a new one
0.12.0
  * Improving Range and Coord limit situations
0.11.0
  * Added package tests
  * Added append/prepend rows/columns to Range
  * Added guess_ods_style function
0.10.0
  * Replaced letter, number parameters by Coord and Range
  * Added compatibilty classes OpenPyXL2010 and ODS_Write_Without_Styles
  * Colors and data styles work in XLSX and ODS
0.9.0
  * Solved problem with freeze and setselectedcell
0.8.0
  * __version__ is now in commons
  * Improved spanish translation
  * Solved problem with merged cells
  * Added overwrite_and_merge to XLSX generator
0.7.0
  * Replaced predefined styles by predefined colors
  * Added officegenerator.xlsx to demo
0.6.0
  * Addapted code to openpyxl-2.4.1
0.5.0
  * [#1] Added dependencies in setup.py
  * [#5] officegenerartor_odf2xml require --file parameter
  * [#6] Show alerts in Windows if can't be executed
  * Added internationalization infrastructure
0.4.0
  * Added pkg_resources support
  * Moved images directory to package
0.3.1
  * Solved bug with image path in demo.py
0.3.0
  * Now officegenerator_demo can delete example files with --remove parameter
  * Added officegenerator_odf2xml to convert odf files to indented xml
0.2.0
  * Added officegenerator_demo to view basic examples and code
0.1.0
  * Basic funcionality

