import unittest
from unittest.mock import MagicMock, patch
import sys
import os
import json

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

# Mock pyhwpx before importing HwpTemplateEngine
sys.modules['pyhwpx'] = MagicMock()
from src.tools.hwp_template_engine import HwpTemplateEngine

class TestHwpTemplateEngine(unittest.TestCase):
    def setUp(self):
        self.mock_hwp_class = sys.modules['pyhwpx'].Hwp
        self.mock_hwp_instance = self.mock_hwp_class.return_value
        self.engine = HwpTemplateEngine(visible=False)

    def test_open_file(self):
        # Setup
        test_path = "test.hwp"
        with patch('os.path.exists', return_value=True):
            # Execute
            result = self.engine.open(test_path)
            
            # Verify
            self.assertTrue(result)
            self.mock_hwp_instance.open.assert_called_once_with(test_path)

    def test_get_field_list(self):
        # Setup
        # pyhwpx.Hwp.GetFieldList returns fields separated by chr(2)
        fields = ["name", "age", "address"]
        field_str = chr(2).join(fields)
        self.mock_hwp_instance.GetFieldList.return_value = field_str
        
        # Execute
        result = self.engine.get_field_list()
        
        # Verify
        self.assertEqual(sorted(result), sorted(fields))

    def test_inject_data(self):
        # Setup
        fields = ["name", "age"]
        self.engine.get_field_list = MagicMock(return_value=fields)
        data = {"name": "John", "age": "30", "unknown": "value"}
        
        # Execute
        count = self.engine.inject_data(data)
        
        # Verify
        self.assertEqual(count, 2)
        self.mock_hwp_instance.put_field_text.assert_any_call("name", "John")
        self.mock_hwp_instance.put_field_text.assert_any_call("age", "30")
        # unknown field should not be called
        with self.assertRaises(AssertionError):
             self.mock_hwp_instance.put_field_text.assert_any_call("unknown", "value")

    def test_process_template(self):
        # Setup
        template_path = "template.hwp"
        output_path = "output.hwp"
        data = {"name": "John"}
        
        with patch('os.path.exists', return_value=True):
            self.engine.get_field_list = MagicMock(return_value=["name"])
            
            # Execute
            result = self.engine.process_template(template_path, data, output_path)
            
            # Verify
            self.assertTrue(result)
            self.mock_hwp_instance.open.assert_called_with(template_path)
            self.mock_hwp_instance.put_field_text.assert_called_with("name", "John")
            self.mock_hwp_instance.save_as.assert_called_with(output_path)

if __name__ == '__main__':
    unittest.main()
