import pdfplumber
import pandas as pd
import numpy as np
from collections import defaultdict

# Configurable parameters
HEADER_FOOTER_KEYWORDS = ["page", "bank", "date", "grand total", "branch", "account", "nomination"]
Y_TOLERANCE_FACTOR = 1.5  # Multiplier of median line height for row clustering
MIN_TEXT_LENGTH = 20  # Skip pages with less text
COLUMN_MERGE_THRESHOLD = 5  # Max gap (pts) to merge adjacent columns

def is_garbled(text):
    """Check if page contains non-table gibberish."""
    clean = ''.join(c for c in text if c.isalnum() or c.isspace()).strip()
    return len(clean) < MIN_TEXT_LENGTH

def extract_tables(pdf_path):
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            try:
                text = page.extract_text()
                if not text or is_garbled(text):
                    continue

                # Calculate dynamic row clustering tolerance
                words = page.extract_words(x_tolerance=2, keep_blank_chars=False)
                if not words:
                    continue
                
                line_heights = [word['bottom'] - word['top'] for word in words]
                median_line_height = np.median(line_heights)
                y_tolerance = median_line_height * Y_TOLERANCE_FACTOR

                # Detect columns using lines or text clustering
                vertical_x = detect_columns(page, words)
                if not vertical_x:
                    continue

                # Cluster rows and split tables on large vertical gaps
                rows = cluster_rows(words, page.height, y_tolerance)
                tables = split_tables(rows, page.height * 0.05)

                for table in tables:
                    table_data = build_table(table, vertical_x)
                    table_data = filter_headers_footers(table_data)
                    if table_data and len(table_data) > 1:
                        all_tables.append((f"Page_{page_num+1}", table_data))
            except Exception as e:
                print(f"Error processing page {page_num+1}: {str(e)}")
                continue
    return all_tables

def detect_columns(page, words):
    """Detect column boundaries using lines or text density."""
    # Try vertical lines first
    vertical_lines = [line for line in page.lines if line['height'] == 0]
    if vertical_lines:
        x_coords = sorted({line['x0'] for line in vertical_lines} | {line['x1'] for line in vertical_lines})
    else:
        # Text-based column detection with merging
        x_coords = [word['x0'] for word in words] + [word['x1'] for word in words]
        if not x_coords:
            return []
        x_coords.sort()
        
        # Merge close x-coordinates
        merged = []
        prev = x_coords[0]
        for x in x_coords[1:]:
            if x - prev <= COLUMN_MERGE_THRESHOLD:
                prev = (prev + x) / 2
            else:
                merged.append(prev)
                prev = x
        merged.append(prev)
        x_coords = merged

    # Add page boundaries
    return [0] + x_coords + [page.width]

def cluster_rows(words, page_height, y_tolerance):
    """Group words into rows with dynamic vertical tolerance."""
    rows = defaultdict(list)
    for word in words:
        base_y = round(word['top'] / y_tolerance) * y_tolerance
        rows[base_y].append(word)
    
    # Sort rows top-to-bottom and split multi-line cells
    return sorted(rows.items(), key=lambda x: x[0])

def split_tables(rows, gap_threshold):
    """Split into separate tables based on vertical gaps."""
    tables = []
    current_table = []
    prev_y = None
    
    for y, words in rows:
        if prev_y is not None and (y - prev_y) > gap_threshold:
            if current_table:
                tables.append(current_table)
                current_table = []
        current_table.append((y, words))
        prev_y = y
    
    if current_table:
        tables.append(current_table)
    return tables

def build_table(table_rows, vertical_x):
    """Organize words into structured table cells."""
    table_data = []
    last_row = None
    
    for y, words in table_rows:
        row = [''] * (len(vertical_x) - 1)
        words_sorted = sorted(words, key=lambda w: w['x0'])
        
        for word in words_sorted:
            for i in range(len(vertical_x) - 1):
                if vertical_x[i] <= word['x0'] < vertical_x[i + 1]:
                    row[i] += f" {word['text']}".strip()
                    break
        
        # Merge with previous row if likely continuation
        if last_row and is_continuation(last_row, row):
            for i in range(len(row)):
                if row[i] and not last_row[i]:
                    last_row[i] = row[i]
        else:
            if last_row:
                table_data.append(last_row)
            last_row = row
    
    if last_row:
        table_data.append(last_row)
    return table_data

def is_continuation(prev_row, curr_row):
    """Check if current row continues previous (multi-line cell)."""
    return sum(1 for p, c in zip(prev_row, curr_row) if c and not p) < 2

def filter_headers_footers(table_data):
    """Remove header/footer rows using keyword matching."""
    return [
        row for row in table_data
        if not any(keyword in ' '.join(row).lower() for keyword in HEADER_FOOTER_KEYWORDS)
    ]

def save_to_excel(tables, output_path):
    with pd.ExcelWriter(output_path) as writer:
        for sheet_name, table in tables:
            df = pd.DataFrame(table[1:], columns=table[0])
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Usage
pdf_path = "your_document.pdf"
output_excel = "output.xlsx"
tables = extract_tables(pdf_path)
if tables:
    save_to_excel(tables, output_excel)
    print(f"Success! Tables saved to {output_excel}")
else:
    print("No tables detected.")
