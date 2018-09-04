## @package officegenerator
## @brief Generate office files

import datetime
from officegenerator.libodfgenerator import ODS_Read,  ODS_Write,  ODT, OdfCell, OdfPercentage, OdfMoney, OdfFormula, OdfSheet,  ODSColumnWidth
from officegenerator.libxlsxgenerator import OpenPyXL
from officegenerator.commons import *

__version__ = '0.7.0'
__versiondate__=datetime.date(2018,9,4)
