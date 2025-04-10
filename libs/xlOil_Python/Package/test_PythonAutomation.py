import unittest
import sys
import os
import numpy as np

from TestConfig import *


class Test_PythonAutomation(unittest.TestCase):
    def __init__(self, methodName: str = "runTest") -> None:
        super().__init__(methodName)

        import xloil
        self._app = xloil.Application()
        self._wb = self._app.workbooks.add()
        # Can be useful for debugging
        # app.visible=True

    def test_range_trim(self):

        import xloil as xlo

        ws = self._wb.worksheets.add()
        address = "A1:C1"
        range = ws[address]

        self.assertEqual(range.address(local=True), address)

        self.assertEqual(len(range.areas), 1)

        # Create the range another way, check we have the same result
        self.assertEqual(range.bounds, 
                         ws.range(from_row=0, from_col=0, 
                                  num_rows=1, num_cols=3).bounds)

        # Empty range: trimming returns top left cell
        firstR, firstC, lastR, lastC = range.trim().bounds
        self.assertEqual(firstR, lastR)
        self.assertEqual(firstC, lastC)

        # Make range non empty, now trimming returns entire range
        range.set(1)
        firstR, firstC, lastR, lastC = range.trim().bounds
        self.assertEqual(range.trim().bounds, range.bounds)

        # Try getting special cells constants - should return
        # the entire range
        numbers = range.special_cells("constants", ("numbers", "logical"))
        self.assertEqual(numbers.bounds, range.bounds)

        # Double values in the range and check the sum doubles
        range *= 2
        self.assertEqual(np.sum(range.value), 2 * len(range))

        # Clear the range, now trimming should give 1 cell again
        range.clear()
        self.assertEqual(range.trim().shape, (1, 1))


    def test_specialcells(self):

        ws = self._wb.worksheets.add()
        ws["A1"] = 1
        ws["A2"] = "world"
        ws["B2"] = 2
        ws["B1"] = "hello"
        
        text_cells = ws["A1:B2"].special_cells("constants", str)

        # Should be a 2 area range
        self.assertEqual(len(text_cells.areas), 2)

        message = " ".join((cell.value for cell in text_cells))
        self.assertEqual(message, "hello world")

    def test_formula(self):
        ws = self._wb.worksheets.add()
        ws["A1"] = "=A2+1"
        ws["A2"] = 2

        self.assertEqual(ws["A1"].has_formula, True)
        self.assertEqual(ws["A1"].formula, "=A2+1")



class Test_PythonUtils(unittest.TestCase):
    
    def test_convert_address_a1(self):
        import xloil as xlo

        self.assertEqual(
            xlo.Address((0, 0))(style="a1"),
            "A1")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).a1,
            "A1:B2")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1)).a1_fixed,
            "$A$1:$B$2")

        self.assertEqual(
            xlo.Address((0, 0, 1, 1))(style="$a$1"),
            "$A$1:$B$2")

        self.assertEqual(
            xlo.Address("Sheet1!r1c1")(style="a1"),
            "'Sheet1'!A1")

        count = 0
        for cell in xlo.Address((0, 0, 2, 2)):
            self.assertEqual(cell.from_row, cell.to_row)
            self.assertEqual(cell.from_col, cell.to_col)
            count += 1

        self.assertEqual(count, 6)

    def test_convert_address_rc(self):
        import xloil as xlo

        self.assertEqual(
            xlo.Address((0, 0, 1, 1))(style="rc"),
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
