import unittest
from officegenerator.commons import Range


## Class to text Range operations. Class must begin with Test and modules with test_ too
class TestRange(unittest.TestCase):
    def test_creation(self):
        s="A1:B2"
        self.assertEqual(Range(s).string(), s)

    def test_appendRow(self):
        s="A1:B2"
        self.assertEqual(Range(s).appendRow(2).string(),"A1:B4")

    def test_appendColumn(self):
        s="A1:B2"
        self.assertEqual(Range(s).appendColumn(2).string(),"A1:D2")

    def test_prependColumn(self):
        s="A1:B2"
        self.assertEqual(Range(s).prependColumn(2).string(),"A1:B2")


if __name__ == '__main__':
    unittest.main()