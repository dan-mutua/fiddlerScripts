# Fiddler Automation Tool

A Python-based automation tool for analyzing Fiddler (.saz) capture files and checking IP ownership across Microsoft and ARIN databases.

## Prerequisites

Before running this tool, ensure you have Python 3.x installed and the following packages:

```bash
pip install pandas
pip install requests
pip install beautifulsoup4
pip install colorama
pip install tkinter
```

## Required Packages

- **pandas**: Data manipulation and Excel file handling
- **requests**: Making HTTP requests to ARIN and Microsoft APIs
- **beautifulsoup4**: Parsing HTML content from Fiddler captures
- **colorama**: Colored console output
- **tkinter**: GUI interface (usually comes with Python)

## Installation

1. Clone or download this repository to your local machine
2. Install the required packages:
```bash
pip install -r requirements.txt
```

## Directory Structure

```
fiddler-automation-tool/
├── main.py           # Main entry point
├── index.py          # Fiddler SAZ file processor
├── searchMsDocs.py   # Microsoft Docs API checker
├── ariSearch.py      # ARIN database checker
└── README.md         # This file
```

## Usage

1. Run the main script:
```bash
python main.py
```

2. When prompted, select your .saz file through the file dialog
3. The tool will automatically:
   - Extract and process the Fiddler capture
   - Check IP addresses against Microsoft Docs
   - Verify IP ownership in ARIN database
   - Generate an Excel report with results

## Output

The tool generates an Excel file (`filtered_results.xlsx`) containing:
- Process names
- IP addresses
- Hostnames
- Microsoft Docs validation results
- ARIN ownership validation results
- Combined pass/fail results

## Troubleshooting

- Ensure your .saz file is not corrupted
- Check write permissions in the script directory
- Verify internet connectivity for API checks

## Error Messages

- "Failed to run [script]": Check individual script logs
- "Error processing IP": Network or API access issue
- "Please provide a .saz file path": No file selected



## Contributors
