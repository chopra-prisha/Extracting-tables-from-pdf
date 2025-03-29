# PDF Table Extractor

## Overview
This script extracts tables from a PDF file and saves them into an Excel file. It uses `pdfplumber` to read and analyze the PDF, detects tables by clustering words into rows and columns, and processes the data to filter out headers and footers before exporting it to Excel.

## Features
- Extracts tabular data from PDF documents.
- Handles multi-line cell merging.
- Removes headers and footers based on predefined keywords.
- Saves extracted tables as separate sheets in an Excel file.

## Dependencies
Ensure you have the following Python libraries installed:
- `pdfplumber`
- `pandas`
- `numpy`

You can install them using:
```sh
pip install pdfplumber pandas numpy
```

## Usage
### Running the script
1. Place your target PDF file in the same directory as the script.
2. Modify the `pdf_path` and `output_excel` variables accordingly.
3. Run the script:
   ```sh
   python script.py
   ```

### Output
- Extracted tables will be saved as an Excel file.
- Each table is stored as a separate sheet within the file.

## Configuration Parameters
- `HEADER_FOOTER_KEYWORDS`: List of keywords to filter out headers and footers.
- `Y_TOLERANCE_FACTOR`: Controls row clustering sensitivity.
- `MIN_TEXT_LENGTH`: Minimum text length to consider a page for table extraction.
- `COLUMN_MERGE_THRESHOLD`: Threshold for merging adjacent detected columns.

## Functions
### `extract_tables(pdf_path)`
Extracts tables from the given PDF file.

### `detect_columns(page, words)`
Detects column boundaries using lines or text clustering.

### `cluster_rows(words, page_height, y_tolerance)`
Groups words into rows based on vertical tolerance.

### `split_tables(rows, gap_threshold)`
Splits content into separate tables based on large vertical gaps.

### `build_table(table_rows, vertical_x)`
Organizes words into structured table cells.

### `filter_headers_footers(table_data)`
Filters out unwanted header and footer rows.

### `save_to_excel(tables, output_path)`
Saves extracted tables to an Excel file.

## Error Handling
The script skips pages that contain too little text or cannot be processed. If an error occurs during extraction, it will print a message indicating the problematic page.

