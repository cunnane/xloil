import unittest
import sys
import os

#
# For unknown reasons, the tests do not run in Visual Studio's test runner
# documentation on python in VS seems very thin so it is unclear how
# the environment should be set up. The test can be invoked at a command 
#


class Test_PythonAutomation(unittest.TestCase):
    def test_A(self):
        # Not great!
        sys.path.append("..\\..\\..\\build\\x64\\Debug")

        import xloil as xlo

        with xlo.Application() as app:
            app.visible=True
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
