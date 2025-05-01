# Southern Starz Inventory Checker

A GUI application to check Southern Starz wine inventory against the company website.

## Features

- User-friendly interface for PDF inventory files
- Drag and drop support (optional)
- Automatically extracts data from PDF inventory reports
- Compares inventory with products on the website
- Checks for product assets (spec sheets, shelf-talkers, etc.)
- Verifies varietal matching between inventory and website
- Highlights varietal mismatches in the output
- Generates Excel reports with inventory status

## Installation

1. Make sure you have Python 3.8+ installed
2. Clone or download this repository
3. Install the required dependencies:

```
pip install -r requirements.txt
```

4. Optional: For drag-and-drop functionality, install tkinterdnd2:
   - Windows: `pip install tkinterdnd2`
   - For macOS/Linux users, you may need to install from source:
     ```
     pip install git+https://github.com/pmgagne/tkinterdnd2
     ```

## Usage

### GUI Application

Run the GUI application:

```
python inventory_gui.py
```

Or use the batch file (Windows):

```
run_inventory_checker.bat
```

1. Select a PDF inventory file by:
   - Clicking the "Browse" button and selecting a file
   - Dragging and dropping a PDF file onto the application (if tkinterdnd2 is installed)

2. Click "Process Inventory" to start the analysis
   - Progress will display in the log window
   - The process may take several minutes depending on the size of the inventory

3. When complete, click "Download Excel" to open the generated report

### Command Line

You can also run the inventory checker from the command line:

```
python inventory_checker.py path/to/inventory.pdf
```

## Notes

- The application requires an internet connection to check the Southern Starz website
- For large inventories, the process may take some time to complete
- The Excel report includes both inventory items and website-only products
- Drag and drop functionality is optional and requires tkinterdnd2 to be properly installed

## Troubleshooting

- If drag and drop doesn't work, you can still use the "Browse" button to select files
- If you encounter issues with tkinterdnd2:
  - Make sure you're using a compatible version
  - Try installing from the GitHub repository as mentioned in the installation section
  - On some systems, you may need additional system libraries

For Windows users, ensure you have Microsoft Visual C++ Redistributable installed for the PDF processing libraries. 