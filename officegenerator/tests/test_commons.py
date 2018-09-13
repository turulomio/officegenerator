import unittest
from officegenerator.commons import Range, Coord


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

class TestCoord(unittest.TestCase):
    def test_appendRow(self):
        s="Z1"
        self.assertEqual(Coord(s).addColumn().string(), "AA1")
        self.assertEqual(Coord(s).addColumn(-25).string(), "A1")
        self.assertEqual(Coord(s).addColumn(-26).string(), "A1")
        self.assertEqual(Coord(s).addRow(-1).string(), "Z1")


if __name__ == '__main__':
    unittest.main()