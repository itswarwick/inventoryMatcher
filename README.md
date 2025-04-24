# Starz Updater

A Python script to automate inventory checking between PDF reports and the Southern Starz website.

## Features

- Extracts inventory data from PDF reports
- Checks product availability on Southern Starz website
- Generates Excel reports with inventory status
- Supports automated web scraping with Selenium

## Setup

1. Clone the repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```
3. Create a `.env` file with your credentials:
```
SOUTHERN_STARZ_USERNAME=your_username
SOUTHERN_STARZ_PASSWORD=your_password
```

## Usage

Run the script with a PDF inventory file:
```bash
python inventory_checker.py path/to/inventory.pdf
```

The script will:
1. Extract inventory data from the PDF
2. Check each product on the Southern Starz website
3. Generate an Excel report with results

## Output

The script generates an `inventory_report.xlsx` file containing:
- SKU
- Producer
- Description
- Quantities (On Hand, On Order, Available)
- Website Status (On Website, Has Bottle Shot, Has Label)
- Product URLs 