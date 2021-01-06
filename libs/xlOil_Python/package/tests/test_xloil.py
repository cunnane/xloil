import unittest
import module_with_converter


class MyTestCase(unittest.TestCase):
    def test_import_module_with_custom_converter(self):
        result = module_with_converter.my_adder(1, 1)  # runs but does not call the custom converter. That is expected.
        self.assertEqual(2, result)


if __name__ == '__main__':
    unittest.main()
