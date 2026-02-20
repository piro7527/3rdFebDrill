---
name: Student Data Aggregation
description: Aggregates student CSV data, summing questions and correct answers per field, and pivoting fields to columns. Also generates per-student HTML/PDF feedback reports.
---

# Student Data Aggregation Skill

This skill allows you to process a collection of Student Learning Record CSV files, aggregate the data by student and field, and export a summarized Excel file.

## Features
- **Merge**: Combines multiple CSV files from a directory.
- **Clean**: Removes duplicate rows.
- **ID Normalization**: Handles cases where a single student has multiple Student IDs by unifying them based on `氏名` (Name). It uses the most frequent (mode) ID for that name.
- **Aggregate**: Sums `問題数` (Questions) and `正答数` (Correct Answers) for each student per `分野` (Field).
- **Pivot**: Arranges the data so that each student is a single row, with fields expanded as columns (e.g., `FieldA_問題数`, `FieldA_正答数`).
- **Exclude Students**: Specific students can be excluded from output (e.g., `EXCLUDE_STUDENTS` list).

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

---

# Drill Feedback Report Generation

The `generate_drill_feedback.py` script generates per-student HTML/PDF feedback reports.

## Features
- **PDF/HTML with Student ID in Filename**: Output files are named `{学籍番号}_{氏名}.pdf` (e.g., `P23112_保田_美咲.pdf`).
- **Name-based Grouping**: Students are grouped by `氏名` (not `学籍番号`), so students with multiple IDs (e.g., `P2205216`/`P2205218` for 瀧口紫音) are automatically merged.
- **Representative ID Selection**: The most frequently used `学籍番号` is chosen as the representative ID for the filename.
- **Exclude Students**: The `exclude_students` parameter filters out specified students (e.g., `["藤野滉大"]`).
- **Period Parameter**: The report period is passed as a string (e.g., `"2026年2月16日〜2月20日"`).

## Usage

```bash
# WeasyPrint requires pango library for PDF generation
# If PDF generation fails, set the library path:
DYLD_FALLBACK_LIBRARY_PATH=$(brew --prefix)/lib python3 generate_drill_feedback.py
```

## Output
- `output/html/` — HTML reports per student (filename: `{学籍番号}_{氏名}.html`)
- `output/pdf/` — PDF reports per student (filename: `{学籍番号}_{氏名}.pdf`)

## Key Configurations (in `__main__` block)
- `csv_path`: Input CSV file (default: `学習記録_統合.csv`)
- `exclude_students`: List of student names to exclude on `extractor.load()`
- `period`: Report period string passed to `report_gen.generate_all()`

## Dependencies
-   `pandas`
-   `openpyxl` (for Excel export)
-   `weasyprint` (for PDF generation)
-   `pango` (system library, install via `brew install pango`)
