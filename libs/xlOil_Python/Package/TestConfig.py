from pathlib import Path
import sys
import os

PACKAGE_PATH = Path(__file__).parent.resolve()
SOLUTION_PATH = PACKAGE_PATH / "../../../"
TEST_PATH = SOLUTION_PATH / "tests"

def _set_test_environment():
    """
        If the env var XLOIL_TEST_BIN_DIR has been defined, ensure the python 
        package is on the path and XLOIL_BIN_DIR is set. Note that 
        XLOIL_TEST_BIN_DIR is assumed to be relative to the solution root.

        Otherwise assume we are in pre-release test mode and that the 
        xloil addin has been installed.
    """

    test_bin_dir = os.environ.get("XLOIL_TEST_BIN_DIR", None)
    
    if test_bin_dir is not None:
        global ADDIN_PATH
        bin_path = (SOLUTION_PATH / test_bin_dir).resolve()

        os.environ['XLOIL_BIN_DIR'] = str(bin_path)
        # Need to do this so when we launch Excel, we can see the python package
        os.environ["PYTHONPATH"] = str(PACKAGE_PATH)
    else:
        os.environ["PYTHONPATH"] = os.curdir

_set_test_environment()