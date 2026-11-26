import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from clean_and_report import clean_amount, clean_category, clean_date, clean_rows


class CleaningTests(unittest.TestCase):
    def test_dates(self):
        self.assertEqual(clean_date("03/01/2026"), "2026-01-03")
        self.assertEqual(clean_date("08-01-2026"), "2026-01-08")
        self.assertIsNone(clean_date("hello"))

    def test_amounts(self):
        self.assertEqual(clean_amount('-4,500.00'), -4500.0)
        self.assertIsNone(clean_amount("abc"))

    def test_category(self):
        self.assertEqual(clean_category("  office   supplies "), "Office Supplies")

    def test_duplicates_removed(self):
        row = {"Date": "2026-01-15", "Description": "Lunch", "Category": "Meals", "Amount": "-430"}
        good, problems = clean_rows([row, dict(row, Category="meals")])
        self.assertEqual(len(good), 1)
        self.assertEqual(problems, [])


if __name__ == "__main__":
    unittest.main()
