import pdfplumber
import pandas as pd
import numpy as np
from sklearn.cluster import DBSCAN
from collections import defaultdict

def extract_tables(pdf_path, excel_path, y_tolerance=5):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=2, y_tolerance=1)
            if not words:
                continue
            
            # Cluster words into logical rows
            y_coords = np.array([word['top'] for word in words]).reshape(-1, 1)
            clustering = DBSCAN(eps=y_tolerance, min_samples=1).fit(y_coords)
            labels = clustering.labels_
            
            # Group words by row cluster and sort vertically
            rows = defaultdict(list)
            for word, label in zip(words, labels):
                rows[label].append(word)
            sorted_rows = sorted(rows.values(), key=lambda r: np.mean([w['top'] for w in r]))
            
            # Detect column boundaries dynamically
            all_x0 = sorted([word['x0'] for row in sorted_rows for word in row])
            hist, bins = np.histogram(all_x0, bins=20)
            peaks = bins[:-1][hist > np.percentile(hist, 70)]
            column_boundaries = np.sort(peaks)
            
            # Extract and merge multi-line rows
            table_data = []
            current_row = []
            for row in sorted_rows:
                row_words = sorted(row, key=lambda w: w['x0'])
                cells = [[] for _ in range(len(column_boundaries) + 1)]
                
                for word in row_words:
                    col = np.digitize(word['x0'], column_boundaries)
                    cells[col].append(word['text'])
                
                # Merge multi-line cells
                if current_row:
                    if len(' '.join(cells[0])) < 2:  # Likely continuation
                        for i, cell in enumerate(cells):
                            if cell:
                                current_row[i] += ' ' + ' '.join(cell)
                    else:
                        table_data.append(current_row)
                        current_row = [' '.join(cell) for cell in cells]
                else:
                    current_row = [' '.join(cell) for cell in cells]
            
            if current_row:
                table_data.append(current_row)
            
            if table_data:
                tables.append(table_data)
    
    # Save to Excel with headers
    with pd.ExcelWriter(excel_path) as writer:
        for i, table in enumerate(tables):
            df = pd.DataFrame(table[1:], columns=table[0])
            df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

# Usage
extract_tables("test6 (1).pdf", "output.xlsx")
