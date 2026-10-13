# Syntecxhub CSV to Excel Converter

A Python command-line tool that converts CSV files into Excel workbooks while cleaning and normalizing data.

## Features

- Convert CSV files to Excel `.xlsx` workbooks.
- Normalize column names to lowercase snake_case.
- Remove extra whitespace from text values.
- Detect and handle missing values.
- Parse date columns and report invalid dates.
- Rename columns using command-line options.
- Create output directories automatically.
- Provide logging and helpful error messages.
- Include automated unit tests.

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Argparse
- Logging
- Unittest

## Project Structure

```text
Syntecxhub_CSV_Excel_Converter/
├── converter.py
├── sample_data.csv
├── requirements.txt
├── README.md
├── .gitignore
└── tests/
    └── test_converter.py
```

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR-USERNAME/Syntecxhub_CSV_Excel_Converter.git
cd Syntecxhub_CSV_Excel_Converter
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it on Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Install the dependencies:

```bash
python -m pip install -r requirements.txt
```

## Usage

Basic conversion:

```bash
python converter.py --input sample_data.csv --output output/cleaned_data.xlsx
```

Rename columns:

```bash
python converter.py --input sample_data.csv --output output/renamed.xlsx --rename "Student Name:Full Name" --rename "Marks:Score"
```

Specify date columns:

```bash
python converter.py --input sample_data.csv --output output/dates.xlsx --date-columns "Join Date"
```

Replace missing values:

```bash
python converter.py --input sample_data.csv --output output/filled.xlsx --fill-missing "N/A"
```

Display help:

```bash
python converter.py --help
```

## Running Tests

```bash
python -m unittest discover -s tests -v
```

All 9 automated tests passed during local testing.

## Example

Input: `sample_data.csv`

Output: `output/cleaned_data.xlsx`

The output workbook contains cleaned column names, trimmed text values, parsed dates, and handled missing values.

## Author

Developed as part of the Syntecxhub Python Programming Internship.

## License

For educational and internship purposes.