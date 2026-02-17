---
name: Student Data Aggregation
description: Aggregates student CSV data, summing questions and correct answers per field, and pivoting fields to columns.
---

# Student Data Aggregation Skill

This skill allows you to process a collection of Student Learning Record CSV files, aggregate the data by student and field, and export a summarized Excel file.

## Features
- **Merge**: Combines multiple CSV files from a directory.
- **Clean**: Removes duplicate rows.
- **ID Normalization**: Handles cases where a single student has multiple Student IDs by unifying them based on `氏名` (Name). It uses the most frequent (mode) or maximum ID for that name.
- **Aggregate**: Sums `問題数` (Questions) and `正答数` (Correct Answers) for each student per `分野` (Field).
- **Pivot**: Arranges the data so that each student is a single row, with fields expanded as columns (e.g., `FieldA_問題数`, `FieldA_正答数`).

## Usage

1.  Place the `organize_students.py` script in the directory containing your CSV files.
2.  Run the script:
    ```bash
    python3 organize_students.py
    ```
3.  The script will generate:
    -   `student_data/`: A folder containing individual CSV files for each student.
    -   `summary_aggregated_[timestamp].xlsx`: A single Excel file with the aggregated data.

## Script Logic (Key Functions)

### `create_aggregated_excel_summary`
This function performs the aggregation and pivoting:
1.  **Load** all CSVs.
2.  **Group** by `['学籍番号', '氏名', '分野']`.
3.  **Sum** numeric columns (`問題数`, `正答数`).
4.  **Pivot** the table to make `分野` the columns.
5.  **Flatten** the column names.
6.  **Export** to Excel.

## Dependencies
-   `pandas`
-   `openpyxl` (for Excel export)
