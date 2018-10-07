## @namespace officegenerator.tests.test_libxlsxgenerator
## @brief Test for officegenerator.libxlsxgenerator functions and classes
import unittest
from officegenerator.libxlsxgenerator import OpenPyXL

## Class to text OpenPyXL methods. Class must begin with Test and modules with test_ too
class TestOpenPyXL(unittest.TestCase):
    ## It creates and removes sheets
    def test_create_remove(self):
         xlsx=OpenPyXL("officegenerator.xlsx")
         xlsx.setCurrentSheet(0)
         xlsx.setSheetName("It was")
         xlsx.createSheet("New")
         xlsx.remove_sheet_by_id(0)
         self.assertEqual(xlsx.sheet_name(0),"New")

if __name__ == '__main__':
    unittest.main()
