# Column Classification and SQL Generation Script

This script processes an Excel file to classify its columns based on predefined database table column names, maps the data to the corresponding database table structure, and generates SQL `INSERT` statements for inserting the data into the database. 

Columns in the Excel file that do not match the database schema are stored in a special column, `other_information`, as a JSON object.

---

## Features

1. Classifies columns in an Excel file to corresponding columns in a database table using semantic similarity (supports French text).
2. Handles unmatched columns by placing their data in a `JSON`-formatted `other_information` column.
3. Generates SQL `INSERT` statements for inserting the Excel data into the database table.
4. Handles non-serializable data types (e.g., `Timestamp`) by converting them to strings.
5. Supports command-line usage via `argparse`.

---

## Prerequisites

Before running the script, ensure the following are installed on your system:

1. **Python 3.8 or later**
2. Required Python libraries:
   - `pandas`
   - `openpyxl`
   - `sentence-transformers`
   - `torch`
   - `json`

Install dependencies using:

```bash
pip install pandas openpyxl sentence-transformers torch


```

## Usage

### Command-Line Arguments

The script is designed to be run from the command line using the following syntax:



```bash
python match_to_excel.py input_excel.xlsx "db_col1,db_col2,db_col3" db_content.json --output_excel matched_output.xlsx

```


- input_excel.xlsx: The Excel file with column headers to match.\\
- "db_col1,db_col2,db_col3": Comma-separated list of database column names.\\
- db_content.json: JSON file containing database content as a list of dictionaries. \\
- --output_excel: (Optional) Path to save the output Excel file. Default is output.xlsx.\\

