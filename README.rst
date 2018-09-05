Source code & Development:
    https://officegenerator.sourceforge.io
OfficeGenerator doxygen documentation:
    http://turulomio.users.sourceforge.net/doxygen/officegenerator/
Web page main developer
    http://turulomio.users.sourceforge.net/

Description
===========
Python module to quickly generate office documents with predefined styles

License
=======
GPL-3

Dependencies
============
  * https://www.python.org/, as the main programming language.
  * https://pypi.org/project/odfpy/, to generate LibreOffice documents.
  * https://pypi.org/project/openpyxl/, to generate MS Office XLSX/XLSM  documents.
  * http://xmlsoft.org/, to execute xmllint in officegenerator_odf2xml

Usage
=====
You can view officegenerator/demo.py to see an example of code

Changelog
=========
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

