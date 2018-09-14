## @namespace officegenerator.tests.test_commons
## @brief Test for officegenerator.commons functions and classes

import unittest
from officegenerator.commons import Range, Coord


## Class to text Range operations. Class must begin with Test and modules with test_ too

class TestCoord(unittest.TestCase):
    def test_Coord_methods(self):
        s="Z1"
        self.assertEqual(Coord(s).addColumn().string(), "AA1")
        self.assertEqual(Coord(s).addColumn(-25).string(), "A1")
        self.assertEqual(Coord(s).addRow(25).string(), "Z26")
        self.assertEqual(Coord(s).letterIndex(), 25)
        self.assertEqual(Coord(s).letterPosition(), 26)
        self.assertEqual(Coord(s).numberIndex(), 0)
        self.assertEqual(Coord(s).numberPosition(), 1)
        self.assertEqual(Coord(s).string(),"Z1")
        self.assertEqual(str(Coord(s)),"Coord <Z1>")

    def test_Coord_methods_in_the_limit(self):
        s="Z1"
        self.assertEqual(Coord(s).addColumn(-26).string(), "A1")
        self.assertEqual(Coord(s).addColumn(-2600).string(), "A1")
        self.assertEqual(Coord(s).addRow(-1).string(), "Z1")

class TestRange(unittest.TestCase):
    def test_Range_methods(self):
        s="A1:B2"
        r="Z10:AA20"
        #Normal use, making range bigger
        self.assertEqual(Range(s).addColumnAfter(2).string(),"A1:D2")
        self.assertEqual(Range(s).addRowAfter(2).string(),"A1:B4")
        self.assertEqual(Range(r).addColumnBefore(2).string(),"X10:AA20")
        self.assertEqual(Range(r).addRowBefore(2).string(),"Z8:AA20")

        self.assertEqual(Range(s).numColumns(), 2)
        self.assertEqual(Range(s).numRows(), 2)
        self.assertEqual(Range(s).string(), s)

    def test_Range_methods_in_the_limit(self):
        s="A1:B2"
        r="Z10:AA20"
        #Negative use, cutting range
        self.assertEqual(Range(s).addColumnAfter(-2).string(),"A1:A2")
        self.assertEqual(Range(s).addRowAfter(-2).string(),"A1:B1")
        self.assertEqual(Range(r).addColumnBefore(-2).string(),"AB10:AA20")
        self.assertEqual(Range(r).addRowBefore(-2).string(),"Z12:AA20")

if __name__ == '__main__':
    unittest.main()