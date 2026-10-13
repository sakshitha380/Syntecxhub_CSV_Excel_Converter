
import tempfile
import unittest
from pathlib import Path

import pandas as pd

from converter import convert_csv, normalize_column_name


class TestCSVExcelConverter(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp_dir.name)
        self.input_file = self.folder / "input.csv"
        self.output_file = self.folder / "output.xlsx"

        pd.DataFrame({
            " Student Name ": [" Ananya ", "Rahul"],
            "Join Date": ["2026-06-10", "invalid-date"],
            "Marks": [89, None],
        }).to_csv(self.input_file, index=False)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_normalize_column_name(self):
        self.assertEqual(
            normalize_column_name(" Student Name "),
            "student_name"
        )

    def test_csv_to_excel_conversion(self):
        result = convert_csv(self.input_file, self.output_file)

        self.assertTrue(self.output_file.exists())
        self.assertEqual(len(result), 2)
        self.assertIn("student_name", result.columns)

    def test_whitespace_is_removed(self):
        result = convert_csv(self.input_file, self.output_file)

        self.assertEqual(result.loc[0, "student_name"], "Ananya")

    def test_invalid_date_is_handled(self):
        result = convert_csv(self.input_file, self.output_file)

        self.assertTrue(pd.isna(result.loc[1, "join_date"]))

    def test_missing_values_are_preserved(self):
        result = convert_csv(self.input_file, self.output_file)

        self.assertTrue(pd.isna(result.loc[1, "marks"]))

    def test_column_renaming(self):
        result = convert_csv(
            self.input_file,
            self.output_file,
            renames={"Student Name": "Full Name"}
        )

        self.assertIn("full_name", result.columns)
        self.assertNotIn("student_name", result.columns)

    def test_missing_file_raises_error(self):
        with self.assertRaises(FileNotFoundError):
            convert_csv(
                self.folder / "missing.csv",
                self.output_file
            )

    def test_wrong_input_extension(self):
        wrong_file = self.folder / "input.txt"
        wrong_file.write_text("hello", encoding="utf-8")

        with self.assertRaises(ValueError):
            convert_csv(wrong_file, self.output_file)

    def test_wrong_output_extension(self):
        with self.assertRaises(ValueError):
            convert_csv(
                self.input_file,
                self.folder / "output.csv"
            )


if __name__ == "__main__":
    unittest.main()
