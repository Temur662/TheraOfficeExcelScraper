import pandas as pd

# --- Configuration ---
# !!! IMPORTANT: Replace these placeholders with your actual values !!!
EXCEL_FILE_PATH = 'jan-may2019(new).XLS' # The name of your Excel file
INSURANCE_COLUMN_NAME = 'Primary Insurance'       # The column header containing the insurance type/payer
CAR_INSURANCE_KEYWORD = 'STATE FARM INSURANCE' # The specific text that identifies car insurance
PATIENT_ID_COLUMN = 'Pat ID'            # The column header for the Patient ID
TOTAL_CHARGE_COLUMN = 'Total'           # The column header for the charge amount
INSURED_ID_COLUMN = 'Insured ID'      # The column for the patient's insurance ID
FIRST_NAME_COLUMN = 'First Name'
LAST_NAME_COLUMN = 'Last Name'
VISIT_DATE_COLUMN = 'Visit'            # The column header for the visit date
PATIENT_BDAY_COLUMN = 'Birthdate' # The column header for the patient's birth date
CPT_COLUMN = 'CPT'
PAYMENT_COLUMN = 'Payment' # Based on image, adjust if needed
ADJUSTMENT_COLUMN = 'Adjust' # Based on image, adjust if needed
CREDITS_COLUMN = 'Credits'
CPT_UNITS = 'Units'
# --- Main Processing ---
# Read the Excel file into a pandas DataFrame
try:
    df = pd.read_excel(EXCEL_FILE_PATH)

    # --- Data Cleaning and Type Conversion ---
    # Convert 'Visit' column to datetime objects, handling potential errors
    # Errors will be turned into NaT (Not a Time)
    df[VISIT_DATE_COLUMN] = pd.to_datetime(df[VISIT_DATE_COLUMN], errors='coerce')

    # Convert 'Total' column to numeric, coercing errors to NaN (Not a Number)
    # and filling NaN with 0 to allow summation.
    df[TOTAL_CHARGE_COLUMN] = pd.to_numeric(df[TOTAL_CHARGE_COLUMN], errors='coerce').fillna(0)

    # Ensure the Insurance column is treated as string for comparison
    df[INSURANCE_COLUMN_NAME] = df[INSURANCE_COLUMN_NAME].astype(str)

    # Ensure other relevant columns are strings
    df[PATIENT_ID_COLUMN] = df[PATIENT_ID_COLUMN].astype(str)
    df[INSURED_ID_COLUMN] = df[INSURED_ID_COLUMN].astype(str)
    df[FIRST_NAME_COLUMN] = df[FIRST_NAME_COLUMN].astype(str)
    df[LAST_NAME_COLUMN] = df[LAST_NAME_COLUMN].astype(str)
    df[PATIENT_BDAY_COLUMN] = pd[PATIENT_BDAY_COLUMN].astype(str)

    # --- Filtering ---
    # Filter the DataFrame to include only rows related to car insurance
    # This checks if the specified keyword is present in the insurance column (case-insensitive)
    # Also filters out rows where the Visit Date could not be parsed (NaT)
    car_insurance_df = df[
        df[INSURANCE_COLUMN_NAME].str.contains(CAR_INSURANCE_KEYWORD, case=False, na=False) &
        df[VISIT_DATE_COLUMN].notna() # Only include rows with valid dates
    ].copy() # Use .copy() to avoid SettingWithCopyWarning

    # Check if any car insurance rows were found
    if car_insurance_df.empty:
        print(f"No valid rows found with '{CAR_INSURANCE_KEYWORD}' in the '{INSURANCE_COLUMN_NAME}' column and valid dates.")
    else:
        # --- Grouping and Aggregation ---
        # Group by Patient ID and aggregate the required information
        patient_summary = car_insurance_df.groupby(PATIENT_ID_COLUMN).agg(
            FirstName=(FIRST_NAME_COLUMN, 'first'), # Get the first occurrence of the first name
            LastName=(LAST_NAME_COLUMN, 'first'),   # Get the first occurrence of the last name
            Birthday=(PATIENT_BDAY_COLUMN, 'first'), # Get the first occurrence of the birth date
            InsuranceProvider=(INSURANCE_COLUMN_NAME, 'first'), # Get the first occurrence of the insurance provider
            InsuranceID=(INSURED_ID_COLUMN, 'first'), # Get the first occurrence of the Insured ID
            FirstServiceDate=(VISIT_DATE_COLUMN, 'min'), # Find the minimum date (earliest)
            LastServiceDate=(VISIT_DATE_COLUMN, 'max'),  # Find the maximum date (latest)
            TotalCharges=(TOTAL_CHARGE_COLUMN, 'sum')    # Sum the total charges
        ).reset_index() # Convert the grouped result back to a DataFrame

        # Format dates for better readability (optional)
        patient_summary['FirstServiceDate'] = patient_summary['FirstServiceDate'].dt.strftime('%m/%d/%Y')
        patient_summary['LastServiceDate'] = patient_summary['LastServiceDate'].dt.strftime('%m/%d/%Y')

        # --- Output Results ---
        print("Summary of Car Insurance Charges per Patient:")
        # Use .to_string() for better console formatting without index
        print(patient_summary.to_string(index=False))

        # Optional: Save the results to a new Excel file
        # output_filename = 'car_insurance_patient_summary.xlsx'
        # patient_summary.to_excel(output_filename, index=False)
        # print(f"\nResults saved to {output_filename}")

except FileNotFoundError:
    print(f"Error: The file '{EXCEL_FILE_PATH}' was not found.")
    print("Please make sure the file name and path are correct")
except KeyError as e:
    print(f"Error: A required column was not found in the Excel file: {e}")
    print("Please ensure all specified column names exist in your file:")
    print(f"  Patient ID: '{PATIENT_ID_COLUMN}'")
    print(f"  Visit Date: '{VISIT_DATE_COLUMN}'")
    print(f"  Total Charge: '{TOTAL_CHARGE_COLUMN}'")
    print(f"  Insurance Payer: '{INSURANCE_COLUMN_NAME}'")
    print(f"  Insured ID: '{INSURED_ID_COLUMN}'")
    print(f"  First Name: '{FIRST_NAME_COLUMN}'")
    print(f"  Last Name: '{LAST_NAME_COLUMN}'")
except Exception as e:
    print(f"An unexpected error occurred: {e}")