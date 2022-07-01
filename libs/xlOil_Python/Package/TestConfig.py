from pathlib import Path
import sys
import os

#
# TODO: how can we detect the acutal build config eg Release/x64?
# 
TEST_PATH = Path(r"..\..\..\tests").resolve()
BIN_PATH  = Path(r"..\..\..\build\x64\Debug").resolve()
PACKAGE_PATH = Path(__file__).parent.resolve()

sys.path.append(str(BIN_PATH))
os.environ["PATH"] += os.pathsep + str(BIN_PATH)

# Need to do this so when we launch Excel, we can see the python package
os.environ["PYTHONPATH"] = str(PACKAGE_PATH)
