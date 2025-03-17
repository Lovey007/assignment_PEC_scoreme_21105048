# ========================================
#  Importing Required Libraries
# ========================================
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from google.colab import files
import os


# ========================================
#  Function to Extract Tables from PDF
# ========================================
def extract_pdf_content(file_path):
    extracted_tables = []
    
    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"Processing page {page_num}...")

            # Extract tables directly if detected
            tables = page.extract_tables()
            for table_index, table in enumerate(tables):
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    extracted_tables.append((f"Page_{page_num}_Table_{table_index + 1}", df))

            # Custom detection for borderless/irregular tables
            words = page.extract_words()  # Extract text with positions
            if words:
                rows = {}
                for word in words:
                    top_pos = round(word['top'], 1)
                    if top_pos not in rows:
                        rows[top_pos] = []
                    rows[top_pos].append(word)

                # Sort rows by vertical position
                sorted_rows = sorted(rows.items())

                # Build DataFrame from detected rows
                detected_table = []
                for _, row_words in sorted_rows:
                    row_text = []
                    last_x0 = None
                    for word in sorted(row_words, key=lambda w: w['x0']):
                        if last_x0 is not None and word['x0'] - last_x0 > 20:  # Threshold for column separation
                            row_text.append("\t")  # Add a tab separator for new column
                        row_text.append(word['text'])
                        last_x0 = word['x1']
                    detected_table.append(row_text)  # Store as a list of columns

                if detected_table:
                    df = pd.DataFrame(detected_table)
                    extracted_tables.append((f"Page_{page_num}_Detected_Table", df))

    return extracted_tables


# ========================================
#  Function to Save Tables to Excel File
# ========================================
def save_tables_to_excel(tables, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if not tables:  
            print("No tables found in the PDF. Adding a placeholder sheet.")
            pd.DataFrame([['No tables were detected in the PDF.']]).to_excel(writer, sheet_name='No_Tables_Found', index=False)
        else:
            for sheet_name, df in tables:
                df.to_excel(writer, sheet_name=sheet_name, index=False)


# ========================================
#  Main Code Execution for Colab
# ========================================

# Install missing packages
!pip install openpyxl  # Ensure openpyxl is installed for saving to Excel

# Upload the PDF file
uploaded = files.upload()
pdf_path = list(uploaded.keys())[0]

output_excel = "extracted_tables.xlsx"

# Step 1: Extract tables from PDF
tables = extract_pdf_content(pdf_path)

# Step 2: Save extracted tables to Excel
save_tables_to_excel(tables, output_excel)

# Step 3: Download the resulting Excel file
files.download(output_excel)

print(f"Tables have been successfully extracted to {output_excel}")
