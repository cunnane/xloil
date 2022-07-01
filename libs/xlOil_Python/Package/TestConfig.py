from pathlib import Path
import sys
import os

PACKAGE_PATH = Path(__file__).parent.resolve()
SOLUTION_PATH = PACKAGE_PATH / "../../../"
TEST_PATH = SOLUTION_PATH / "tests"
ADDIN_PATH = None


def _set_test_environment():
    """
        If the env var XLOIL_TEST_BIN_DIR has been defined, set paths
        so that the xlOil 'pyd' and 'xll' are loaded from that directory.
        Also ensure the python package is on the path.

        Otherwise assume we are in pre-release test mode and that the 
        xloil addin has been installed.
    """

    bin_dir = os.environ.get("XLOIL_TEST_BIN_DIR", None)
    
    if bin_dir is not None:
        global ADDIN_PATH
        bin_path = (SOLUTION_PATH / bin_dir).resolve()

        # Need to do this so when we launch Excel, we can see the python package
        os.environ["PYTHONPATH"] = str(PACKAGE_PATH)

        sys.path.append(str(bin_path))
        os.environ["PATH"] += os.pathsep + str(bin_path)

        ADDIN_PATH = str(bin_path / "xloil.xll")

_set_test_environment()