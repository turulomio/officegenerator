## @package officegenerator
## @brief Generate office files with predefined styles
from officegenerator.libodfgenerator import ODSStyleColor, ODSStyleCurrency, ColumnWidthODS, ODS_Read, ODS_Write, ODT, ODT_Standard, ODT_Manual_Styles, OdfCell, OdfSheet, guess_ods_style
from officegenerator.libxlsxgenerator import OpenPyXL, ColumnWidthXLSX
from officegenerator.commons import Currency, Percentage, columnAdd, rowAdd, makedirs, Coord, Range, __version__, __versiondate__
