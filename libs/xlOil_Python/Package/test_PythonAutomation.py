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
class Test_PythonUtils(unittest.TestCase):
    
    def test_convert_address_a1(self):
        import xloil as xlo

        self.assertEqual(
            xlo.Address((0, 0)).string(style="a1"),
            "A1")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).a1,
            "A1:B2")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).a1_fixed,
            "$A$1:$B$2")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).string(style="$a$1"),
            "$A$1:$B$2")

        self.assertEqual(
            xlo.Address("Sheet1!r1c1").string(style="a1"),
            "'Sheet1'!A1")

        count = 0
        for cell in xlo.Address((0, 0, 2, 2)):
            self.assertEqual(cell.from_row, cell.to_row)
            self.assertEqual(cell.from_col, cell.to_col)
            count += 1
        self.assertEqual(count, 4)

    def test_convert_address_rc(self):
        import xloil as xlo

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).string(style="rc"),
            "R1C1:R2C2")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).rc_fixed,
            "$R1$C1:$R2$C2")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).rc_fixed,
            "$R1$C1:$R2$C2")

        self.assertEqual(
            xlo.Address("Sheet1!$a$1").rc,
            "'Sheet1'!R1C1")

    def test_convert_address_tuple(self):
        import xloil as xlo

        self.assertEqual(
            xlo.Address((0, 0)).tuple,
            (0, 0))

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).tuple,
            (0, 0, 1, 1))

        self.assertEqual(
            xlo.Address("$a$1").tuple,
            (0, 0))

if __name__ == '__main__':
    unittest.main()
