## @namespace officegenerator.tests.test_demo
## @brief Test for officegenerator_demo script

import unittest
from officegenerator.demo import main


## Class to text Range operations. Class must begin with Test and modules with test_ too

class TestArgs(unittest.TestCase):
    def test_Coord_methods(self):
        main(["--create",])
        main(["--remove",])


if __name__ == '__main__':
    unittest.main()