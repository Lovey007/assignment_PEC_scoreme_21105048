# assignment_PEC_scoreme_21105048

## PDF Table Extractor

A tool designed to detect and extract tables from system-generated PDFs without using Tabula, Camelot, or image conversion. The extracted tables are exported to Excel sheets, handling both bordered, borderless, and irregular-shaped tables.

## Features

Extracts tables from system-generated PDFs with or without borders.
Handles tables with irregular shapes and structures.
Saves extracted tables to Excel files with proper formatting.

## Requirements

### The tool uses the following Python libraries:
pdfplumber - For extracting text and table structures from PDFs.
pandas - For handling and exporting table data.
openpyxl - For writing extracted tables to Excel files.

### Install dependencies via:
!pip install pdfplumber pandas openpyxl

## Usage

### 1. Run on Google Colab
Upload the pdfTableExtractor.py file to your Google Colab environment.
Upload your PDF file when prompted.
Download the resulting Excel file once the extraction is complete.

### 2. Run Locally (Optional)
If you prefer to run the tool locally, make sure you have the necessary packages installed:
pip install pdfplumber pandas openpyxl
Then, execute the script:
python pdfTableExtractor.py

## Code Explanation

### extract_pdf_content(file_path)
Reads the PDF and attempts to extract tables using pdfplumber.
Uses custom detection logic to handle borderless and irregular tables.
Tables are saved as DataFrames.

### save_tables_to_excel(tables, output_file)
Stores the extracted tables into an Excel file.
Handles cases where no tables are detected by adding a placeholder sheet.

### Main Execution
Prompts user to upload a PDF file.

Extracts tables and downloads the resulting Excel file.

## Example Output

### The extracted tables are saved as an Excel file with multiple sheets named:
Page_X_Table_Y for detected tables by pdfplumber.
Page_X_Detected_Table for custom-extracted borderless/irregular tables.

## Future Enhancements

Improve column detection logic for more complex structures.
Support for extracting multi-line cells.
Handling of merged cells and irregular-shaped tables.
