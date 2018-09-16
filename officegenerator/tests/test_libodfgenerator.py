## @namespace officegenerator.tests.test_libodfgenerator
## @brief Test for officegenerator.libodfgenerator functions and classes
import datetime
import unittest
from decimal import Decimal
from officegenerator.commons import Percentage, Currency
from officegenerator.demo import demo_ods
from officegenerator.libodfgenerator import ODS_Read, guess_ods_style

## Class to text Range operations. Class must begin with Test and modules with test_ too
class TestFunctions(unittest.TestCase):
    def test_guess_ods_style(self):
        s=Percentage(1,2)
        self.assertEqual(guess_ods_style("Orange", s), "OrangePercentage")

## Class to text ODS_Read operations. Class must begin with Test and modules with test_ too
class TestODS_Read(unittest.TestCase):
    def test_guess_ods_style(self):
        demo_ods()
        doc=ODS_Read("officegenerator.ods")
        s1=doc.getSheetElementByIndex(0)
        self.assertEqual(doc.getCellValue(s1,"A2"),"Percentage")
        #self.assertEqual(doc.getCellValue(s1,"B2").value, Percentage(21.43,100).value)
        self.assertEqual(doc.getCellValue(s1,"B4").upper(),"=SUM(B2:B3)")
        self.assertEqual(doc.getCellValue(s1,"B6"), Decimal("100.26"))

        s3=doc.getSheetElementByIndex(2)
        self.assertEqual(doc.getCellValue(s3,"B2").__class__, datetime.datetime)
        self.assertEqual(doc.getCellValue(s3,"C2").__class__, datetime.date)
        self.assertEqual(doc.getCellValue(s3,"E2").string(), Currency(12.56, "EUR").string())

if __name__ == '__main__':
    unittest.main()
