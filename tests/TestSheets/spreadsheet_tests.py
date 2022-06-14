from argparse import ArgumentParser
import os
import sys
from pathlib import Path

RESULT_RANGE_PREFIX = "test_"

#
# Process our cmd line args
#
parser = ArgumentParser()
parser.add_argument("--bin", help="path to xloil")
parser.add_argument("testdir", help="path to tests")
args = parser.parse_args()

bin_path = args.bin

sys.path.append(bin_path)
os.environ["PATH"] += os.pathsep + bin_path

import xloil as xlo

test_sheets = [(Path(args.testdir) / x).resolve() for x in Path(args.testdir).glob("*.xlsx")]

app = xlo.Application()
app.visible = True

app.workbooks().add() # Cannot get com object without this

app.to_com().Visible = True

# Load addin
if not app.to_com().RegisterXLL(os.path.join(bin_path, "xloil.xll")):
    raise Exception("xloil load failed")

#input("Attach debugger now...")

test_results = {}
for filename in test_sheets:
    #with app.open(str(filename)) as wb:
    print(filename)
    wb = app.open(str(filename), read_only=True)
    app.calculate(full=True)
    for named_range in wb.to_com().Names:
        if named_range.Name.lower().startswith(RESULT_RANGE_PREFIX):
            # skip one char as RefersTo always starts with '='
            address = named_range.RefersTo[1:]
            test_results[(filename.stem, named_range.Name)] = wb[address].value

    wb.close(save=False)
    del wb


for addin in app.to_com().AddIns:
    addin.Installed = False
    

app.quit()


for k,v in test_results.items():
    print(k, v)

if not all(test_results.values()):
    print("-->FAILED<--")

    