import unittest
import sys
import os

from TestConfig import *


class Test_PythonAutomation(unittest.TestCase):
    def test_A(self):

        import xloil as xlo

        with xlo.Application() as app:

            # Can be useful for debugging
            # app.visible=True
            
            wb = app.workbooks.add()
            ws = wb.worksheets.add()
            range = ws["A1:C1"]

            # Empty range: trimming returns top left cell
            firstR, firstC, lastR, lastC = range.trim().bounds
            self.assertEqual(firstR, lastR)
            self.assertEqual(firstC, lastC)

            # Make range non empty
            range.set(1)
            firstR, firstC, lastR, lastC = range.trim().bounds
            self.assertEqual(firstR, 0)
            self.assertEqual(firstC, 0)
            self.assertEqual(lastR, 0)
            self.assertEqual(lastC, 2)


if __name__ == '__main__':
    input("Attach debugger now...")
    unittest.main()
