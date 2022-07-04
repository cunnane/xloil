import unittest
import os
import sys
from pathlib import Path

from TestConfig import *

RESULT_RANGE_PREFIX = "test_"

# TODO: as new sheets are unlikely to be added often, could manually make a test case per sheet
class Test_SpreadsheetRunner(unittest.TestCase):
    def test_RunSheets(self):

        SHEET_PATH = TEST_PATH / "AutoSheets"

        import xloil as xlo
        from xloil._paths import ADDIN_NAME, XLOIL_BIN_DIR

        test_sheets = [(SHEET_PATH / x) for x in SHEET_PATH.glob("*.xls*")]

        app = xlo.Application()
        app.visible = True

        # Load addin: when running via COM automation, no addins are  
        # loaded by default
        if not app.RegisterXLL(os.path.join(XLOIL_BIN_DIR, ADDIN_NAME)):
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

        
        # Subtests are broken in VS code up to 2021. Don't seem to work
        # in Visual Studio 2019 either... hoping that one day....
        # https://github.com/microsoft/vscode-python/issues/17561

        succeed = True
        for k, v in test_results.items():
            with self.subTest(msg=k):
                self.assertTrue(v)
            if v is not True:
                print(k, v)
                succeed = False
        self.assertTrue(succeed) # Required because VS is broken


if __name__ == '__main__':
    unittest.main()
