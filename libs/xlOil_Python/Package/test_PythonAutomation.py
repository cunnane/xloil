import unittest
import sys

class Test_PythonAutomation(unittest.TestCase):
    def test_A(self):

        sys.path.append("..\\..\\..\\build\\x64\\Debug")
        import xloil as xlo

        with xlo.Application() as app:
            app.visible=True
            wb = app.workbooks.add()
            ws = wb.worksheets.add()
            #ws = wb.add()
            #rng = ws["A1:C1"]

            #rng.trim().bounds
            #rng.set(1)
            #rng.trim().bounds

            #self.assertTrue(True)
            #self.assertEqual(1, 1)


if __name__ == '__main__':
    unittest.main()
