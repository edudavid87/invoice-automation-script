# Invoice Automation with Python

This Python script automates the process of updating values across multiple Excel-based Sales Order (SO) files using a centralized source file. It was originally developed to streamline a repetitive manual workflow in the finance operations team, improving accuracy and saving time.

## Use Case

Finance and operations teams often need to reconcile values—such as revenue, taxes, or balances—across dozens of spreadsheets. This script reads from a master transactions workbook and populates specified cells in various SO files, reducing manual data entry and ensuring consistency.

## Key Features

- Reads input from a structured source workbook (e.g., `Transactions.xlsx`)
- Updates specific cells in a list of target Excel files
- Automates repetitive Excel-based tasks with error handling and logs
- Reduces human error and saves hours of manual processing

## Technologies Used

- Python 3.x
- `openpyxl` – for reading and writing Excel files

## 📂 File Structure
