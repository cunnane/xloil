import unittest
import sys
import os

from TestConfig import *


class Test_PythonAutomation(unittest.TestCase):
    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)

        import xloil
        self._app = xloil.Application()
        self._wb = self._app.workbooks.add()

    def test_A(self):

        import xloil as xlo

        # Can be useful for debugging
        # app.visible=True
            
        ws = self._wb.worksheets.add()
        range = ws["A1:C1"]

        # Empty range: trimming returns top left cell
        firstR, firstC, lastR, lastC = range.trim().bounds
        self.assertEqual(firstR, lastR)
        self.assertEqual(firstC, lastC)

        # Make range non empty, now trimming returns entire range
        range.set(1)
        firstR, firstC, lastR, lastC = range.trim().bounds
        self.assertEqual(range.trim().bounds, range.bounds)

        numbers = range.special_cells("constants", "numbers")
        self.assertEqual(numbers.bounds, range.bounds)


    def test_specialcells(self):

        ws = self._wb.worksheets.add()
        ws["A1"] = 1
        ws["A2"] = "world"
        ws["B2"] = 2
        ws["B1"] = "hello"
        
        text_cells = ws["A1:B2"].special_cells("constants", str)
        message = " ".join((cell.value for cell in text_cells))
        self.assertEqual(message, "hello world")


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
