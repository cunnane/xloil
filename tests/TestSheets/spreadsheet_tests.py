from argparse import ArgumentParser
import os
import sys
from pathlib import Path

RESULT_RANGE_PREFIX = "test_"

#
# Process our cmd line args
#
parser = ArgumentParser(description="Typical usage: spreadsheet_tests.py --bin=..\\..\\build\\x64\\Debug .")
parser.add_argument("--bin", help="path to xloil")
parser.add_argument("testdir", help="path to tests")
args = parser.parse_args()

bin_path = args.bin

sys.path.append(bin_path)
os.environ["PATH"] += os.pathsep + bin_path

import xloil as xlo

test_sheets = [(Path(args.testdir) / x).resolve() for x in Path(args.testdir).glob("*.xls*")]

app = xlo.Application()
app.visible = True

app.workbooks().add() # Cannot get com object without this

# Load addin
if not app.to_com().RegisterXLL(os.path.join(bin_path, "xloil.xll")):
    raise Exception("xloil load failed")

# Uncomment this to pause so the debugger can be attached to the 
# Excel or python processes
#input("Attach debugger now...")
test_results = {}
for filename in test_sheets:
    #with app.open(str(filename)) as wb:
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

for k,v in test_results.items():
    print(k, v)

if not all(test_results.values()):
    print("-->FAILED<--")

    