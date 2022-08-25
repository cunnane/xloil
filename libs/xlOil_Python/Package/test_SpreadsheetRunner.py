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
        
        # Excel like to have at least one workbook open. If not, then when
        # we close our test workbook below, sometimes the COM server will
        # become unavailable and we'll get "Remote procedure call failed"
        dummy = app.workbooks.add()

        # Load addin: when running via COM automation, no addins are  
        # loaded by default
        if not app.RegisterXLL(os.path.join(XLOIL_BIN_DIR, ADDIN_NAME)):
            raise Exception("xloil load failed")

        # Uncomment these lines to help debugging. Note the debugger
        # can be attached to the Excel *or* python process
        #input("Attach debugger now...")
        app.visible = True

        test_results = {}
        for filename in test_sheets:
            print(filename)
            wb = app.open(str(filename), read_only=True)
    
            app.calculate(full=True)
            names = wb.to_com().Names
    
            # Some of the python test functions (not RTD or async) require this 
            # wait time to work. I'm not completely sure why.
            if "settings_wait" in [x.Name.lower() for x in names]:
                wait_time = wb["Settings_Wait"].value
                import time
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
