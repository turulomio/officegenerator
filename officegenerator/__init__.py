## @package officegenerator
## @brief Generate office files with predefined styles
from officegenerator.libodfgenerator import Color, ColumnWidthODS, ODS_Read, ODS_Write, ODT, OdfCell, OdfSheet, guess_ods_style
from officegenerator.libxlsxgenerator import OpenPyXL, ColumnWidthXLSX
from officegenerator.commons import Currency, Percentage, columnAdd, rowAdd, makedirs, Coord, Range
