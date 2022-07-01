import unittest
import os
import sys
from pathlib import Path

from TestConfig import *

RESULT_RANGE_PREFIX = "test_"

class Test_SpreadsheetRunner(unittest.TestCase):
    def test_RunSheets(self):

        SHEET_PATH = TEST_PATH / "AutoSheets"

        import xloil as xlo

        test_sheets = [(SHEET_PATH / x) for x in SHEET_PATH.glob("*.xls*")]

        app = xlo.Application()
        app.visible = True

        # Load addin
        if not app.RegisterXLL(str(BIN_PATH / "xloil.xll")):
            raise Exception("xloil load failed")

        # Uncomment this to pause so the debugger can be attached to the 
        # Excel or python processes
        #input("Attach debugger now...")
        test_results = {}
        for filename in test_sheets:
            print(filename)
            wb = app.open(str(filename), read_only=True)
    
            app.calculate(full=True)
            names = wb.to_com().Names
    
            if "settings_wait" in [x.Name.lower() for x in names]:
                wait_time = wb["Settings_Wait"].value
                import time
                app.calculate()
                time.sleep(wait_time)
                app.calculate()
        
            for named_range in names:
                if named_range.Name.lower().startswith(RESULT_RANGE_PREFIX):
                    # skip one char as RefersTo always starts with '='
                    address = named_range.RefersTo[1:]
                    test_results[(filename.stem, named_range.Name)] = wb[address].value
        
            wb.close(save=False)

        app.quit()

        for k, v in test_results.items():
            with self.subTest(msg=k):
                self.assertEqual(v, True)

if __name__ == '__main__':
    unittest.main()
