# PDF to Excel Converter

A Python tool that extracts tables from PDF files and converts them to Excel format with data cleaning and analysis capabilities.

## Features

- **PDF Table Extraction**: Automatically detects and extracts tables from PDF documents
- **Multi-page Support**: Processes all pages in the PDF to find tables
- **Data Cleaning**: Removes empty rows and columns for cleaner output
- **Excel Export**: Saves all extracted tables to separate Excel sheets
- **Data Analysis**: Calculates column-wise means for numeric data
- **Error Handling**: Gracefully handles non-numeric data and mixed content types

## Requirements

- Python 3.6+
- pdfplumber
- pandas
- numpy
- openpyxl

## Installation

1. Clone or download this project
2. Install the required dependencies:

```bash
pip install pdfplumber pandas numpy openpyxl
```

## Usage

1. Place your PDF file in the project directory (currently configured for `data2.pdf`)
2. Run the converter:

```bash
python canverter.py
```

3. The tool will:
   - Extract all tables from the PDF
   - Clean the data by removing empty rows/columns
   - Calculate statistics for numeric columns
   - Save all tables to `output.xlsx` with separate sheets

## Output

- **Console Output**: Shows processing progress, table counts, and statistics
- **Excel File**: `output.xlsx` containing all extracted tables in separate sheets
- **Sheet Naming**: Each table gets its own sheet named `Table_1`, `Table_2`, etc.

## Configuration

To use a different PDF file, modify the `pdf_path` variable in `canverter.py`:

```python
pdf_path = "your_pdf_file.pdf"
```

## Features in Detail

### Table Detection
- Automatically identifies table structures in PDF pages
- Handles multiple tables per page
- Supports various table formats and layouts

### Data Processing
- Converts extracted data to pandas DataFrames
- Removes completely empty rows and columns
- Converts numeric data where possible
- Calculates column-wise means for numeric columns

### Excel Export
- Creates separate sheets for each table
- Preserves original table structure
- Excludes index and header columns for clean output

## Error Handling

The tool includes robust error handling for:
- Non-numeric data in tables
- Mixed data types
- Empty or malformed tables
- PDF processing errors

## Example Output

```
Page 1: 2 tables found
NumPy array shape: (15, 8)
Column-wise means: [25.5, 30.2, 45.1, 12.8, 67.3, 23.4, 18.9, 34.6]
Writing table to sheet: Table_1, shape: (15, 8)
Total tables found in PDF: 2
```

## Limitations

- Requires well-structured tables in the PDF
- May not handle complex table layouts perfectly
- Numeric calculations only work on columns with consistent data types

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## License

This project is open source and available under the MIT License.
