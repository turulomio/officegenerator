## @namespace officegenerator.tests.test_libxlsxgenerator
## @brief Test for officegenerator.libxlsxgenerator functions and classes
import unittest
from officegenerator.libxlsxgenerator import XLSX_Write

## Class to text XLSX_Write methods. Class must begin with Test and modules with test_ too
class TestXLSX_Write(unittest.TestCase):
    ## It creates and removes sheets
    def test_create_remove(self):
         xlsx=XLSX_Write("officegenerator.xlsx")
         xlsx.setCurrentSheet(0)
         xlsx.setSheetName("It was")
         xlsx.createSheet("New")
         xlsx.remove_sheet_by_id(0)
         self.assertEqual(xlsx.sheet_name(0),"New")

if __name__ == '__main__':
    unittest.main()
