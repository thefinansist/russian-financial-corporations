import importlib.util
import sys
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "update_financials_open.py"
SPEC = importlib.util.spec_from_file_location("financials_updater", MODULE_PATH)
UPDATER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = UPDATER
SPEC.loader.exec_module(UPDATER)


class ParserTests(unittest.TestCase):
    def test_clean_number(self):
        self.assertEqual(UPDATER.clean_number("1 869"), 1869.0)
        self.assertEqual(UPDATER.clean_number("(12,5%)"), -12.5)
        self.assertIsNone(UPDATER.clean_number("—"))

    def test_extracts_last_ltm_value(self):
        html = """
        <table>
          <tr field="net_income">
            <td>10.0</td><td>12.0</td><td class="ltm_spc">&nbsp;</td><td>45.5</td>
          </tr>
          <tr field="roe">
            <td>18.0%</td><td class="ltm_spc">&nbsp;</td><td>21.4%</td>
          </tr>
        </table>
        """
        fields = UPDATER.extract_smartlab_fields(html)
        self.assertEqual(fields["net_income"], 45.5)
        self.assertEqual(fields["roe"], 21.4)


if __name__ == "__main__":
    unittest.main()
