# TheraOffice Excel Scraper

## Overview

This Python script automates the process of generating Personal Injury Protection (PIP) demand letters for patients involved in auto accidents. It reads patient and billing data from an Excel spreadsheet, processes the information using the `pandas` library, identifies relevant auto insurance claims, calculates outstanding balances, and generates individual PDF demand letters for each qualifying patient using the `reportlab` library via a custom `GenPDF` module.

This project was developed in collaboration with a law firm to streamline the creation of over 2,000 demand letters.

## Features

* Reads patient and billing data directly from an Excel file (`.XLS` or `.xlsx`).
* Cleans and formats data, handling potential errors in date and numeric columns.
* Filters records specifically for patients with claims under designated auto insurance providers.
* Calculates the outstanding balance for each service line and aggregates the total balance per patient.
* Groups charges and patient details by Patient ID.
* Generates individual, formatted PDF demand letters for each patient using `reportlab`.
* Provides console output summarizing the processed information for each patient.
* Includes basic error handling for file not found and missing columns.

## How it Works

1.  **Configuration:** Sets up constants for the input Excel file path and the names of relevant columns (e.g., Insurance Provider, Patient ID, Visit Date, Charges, Payments, etc.). It also defines a list of keywords (`AUTO_INSURANCES`) to identify auto insurance providers.
2.  **Data Loading:** Reads the specified Excel file into a `pandas` DataFrame.
3.  **Data Cleaning:**
    * Converts date columns to datetime objects.
    * Converts financial columns (Total Charges, Payments, Adjustments, Credits) to numeric types, handling errors and filling missing values with 0.
    * Ensures other key columns are treated as strings.
4.  **Balance Calculation:** Computes the `LineBalance` for each individual charge line by subtracting payments, adjustments, and credits from the total charge.
5.  **Filtering:** Selects rows where the `Primary Insurance` column contains one of the keywords from the `AUTO_INSURANCES` list and the visit date is valid.
6.  **Grouping & Aggregation:**
    * Groups the filtered data by `Pat ID`.
    * Aggregates summary information for each patient (Name, Insurance ID, Dates of Service, Total Balance, etc.).
    * Collects detailed charge information (Date, CPT, Provider, Facility, Line Balance, etc.) into a list for each patient.
7.  **PDF Generation:** Iterates through the summarized patient data. For each patient, it calls the `generate_pdf` function (from the `GenPDF` module, which utilizes the `reportlab` library) to create a PDF demand letter named `FirstName_LastName.pdf`.
8.  **Output:** Prints a summary of each patient's information and charge details to the console after generating their PDF.
9.  **Error Handling:** Catches `FileNotFoundError` if the Excel file doesn't exist and `KeyError` if expected columns are missing, providing informative messages.

## Configuration

Before running the script, you need to configure the constants at the beginning of the Python file (`.py`):

* `EXCEL_FILE_PATH`: The path to your input Excel file (e.g., `'JanMay2019Final.XLS'`).
* `AUTO_INSURANCES`: A Python list of strings containing names or keywords to identify auto insurance companies in the insurance column. The script uses case-insensitive matching.
* **Column Names:** Update the variables like `INSURANCE_COLUMN_NAME`, `PATIENT_ID_COLUMN`, `TOTAL_CHARGE_COLUMN`, `VISIT_DATE_COLUMN`, etc., to match the exact column headers in *your* Excel file.

```python
# --- Configuration ---
EXCEL_FILE_PATH = 'Your_Input_Data.xlsx' # CHANGE THIS
INSURANCE_COLUMN_NAME = 'Primary Insurance'       # Adjust if your column name is different
PATIENT_ID_COLUMN = 'Pat ID'            # Adjust if your column name is different
# ... other column name constants ...
AUTO_INSURANCES = [ 'STATE FARM', 'GEICO', ... ] # Add/Remove insurers as needed
