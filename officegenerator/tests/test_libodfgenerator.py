import unittest
from officegenerator.commons import Range, Percentage
from officegenerator.libodfgenerator import guess_ods_style


## Class to text Range operations. Class must begin with Test and modules with test_ too
class TestRange(unittest.TestCase):
    def test_guess_ods_style(self):
        s=Percentage(1,2)
        self.assertEqual(guess_ods_style("Orange", s), "OrangePercentage")


if __name__ == '__main__':
    unittest.main()