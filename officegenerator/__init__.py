## @package officegenerator
## @brief Generate office files with predefined styles
from officegenerator.libodfgenerator import ODSStyleColor, ODSStyleCurrency, ColumnWidthODS, ODS_Read, ODS_Write, ODT, ODT_Standard, ODT_Manual_Styles, OdfCell, OdfSheet, guess_ods_style
from officegenerator.libxlsxgenerator import XLSX_Write, XLSX_Read, ColumnWidthXLSX
from officegenerator.commons import columnAdd, rowAdd, Coord, Range, __version__, __versiondate__, __versiondatetime__
from officegenerator.standard_sheets import Model, Model_Auto
from officegenerator.objects.currency import Currency
from officegenerator.objects.percentage import Percentage
