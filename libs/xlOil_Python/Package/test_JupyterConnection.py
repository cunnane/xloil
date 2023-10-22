import unittest
from pathlib import Path
from unittest.mock import Mock
from TestConfig import *


def _test_func1():
    ...

def _test_func2(x):
    ...

def _test_func3(x: str, y:int, **kwargs):
    ...

class Test_JupyterConnection(unittest.TestCase):

    def test_RegisterFunc(self):
        """
        Tests the serialisation and deserialisation which occurs when a function
        is decorated for registration in a jupyter session
        """
        from xloil.jupyter import _JupyterConnection

        # Mock xlOil's function registration as it only works when embedded
        import xloil_core
        registrar = Mock()
        xloil_core._register_functions = registrar

        # Create the impl object which gets created in the jupyter kernel.
        from xloil.jupyter_kernel import _xlOilJupyterImpl
        ipy_mock = Mock()
        xloil_jupyter = _xlOilJupyterImpl(ipy_shell=ipy_mock, excel_hwnd=0)

        # func_inspect.Arg gets pasted into the impl object when created in 
        # a jupyter kernel; here we need to do that manually
        from xloil.func_inspect import Arg
        setattr(_xlOilJupyterImpl, "Arg", Arg)
        
        for func in [_test_func1, _test_func2, _test_func3]:
            # Subtests are broken in Visual Studio as of 2022

            #with self.subTest(func=func):

            # 'decorate' func using the decorator exposed in jupyter
            xloil_jupyter.func(func)

            # Get the args the decorator publishes to kernels
            args = ipy_mock.display_pub.publish.call_args[0]

            self.assertEqual(args[1].get('type'), "FuncRegister")

            func_info = args[0].get('xloil/data')

            _JupyterConnection._process_xloil_message(
                self=Mock(), message_type="FuncRegister", payload=func_info)

            registrar.assert_called()


