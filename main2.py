import pandas as pd
from GenPDF import generate_pdf
AUTO_INSURANCES = [
    'STATE FARM INSURANCE',
    'GEICO',
    'PROGRESSIVE',
    'ALLSTATE',
    'NATIONWIDE',
    'USAA',
    'LIBERTY MUTUAL',
    'GEICO INSURANCE',
    'United Auto Insurance Company',
    'Progressive Insurance',
    'Progressive Auto Insurance'
    'USAA Auto Insurance',
    'Progressive Insurance Company',
    'Allstate Insurance',
    'Allstate Insurance Company',
    'INFINITY AUTO INS',
    'Imperial Fire & Casualty Insurance',
    'Hartford Insurance',
    'Liberty Mutual Insurance Company',
    'National General Insurance Company',
    "Great West Casualty Company",
    "OLD AMERICAN INDEMNITY COMPANY",
    "Bristol West Insurance",
    "Bristol West Insurance Company",
    "Security National Insurance Company",
    "Security National Life"
    "Kemper "
    "UNITED AUTO",
    "United Auto Insurance Company",
    "Falcon Insurance Group",
    "Security Insurance Company"
    "GEICO",
    "STATE FARM",
    "PROGRESSIVE",
    "ALLSTATE",
    "Responsive Auto Insurance",
    "USAA",
    "LIBERTY MUTUAL",
    "FARMERS",
    "NATIONWIDE",
    "AMERICAN FAMILY",
    "TRAVELERS",
    "METLIFE",
    "ESURANCE",
    "THE HARTFORD",
    "AAA",
    "SAFECO",
    "MERCURY INSURANCE",
    "ERIE INSURANCE",
    "COUNTRY FINANCIAL",
    "NJM INSURANCE",
    "AUTO-OWNERS INSURANCE",
    "AUTO OWNERS INSURANCE",
    "CSAA INSURANCE GROUP",
    "KEMPER",
    "Responsive Auto Insurance",
    "MAPFRE INSURANCE",
    "ROOT INSURANCE",
    "DISTRIBUTED INSURANCE",
    "CLEARCOVER",
    "BRANCH INSURANCE",
    "THE GENERAL",
    "DAIRYLAND INSURANCE",
    "GAINSCO",
    "INFINITY INSURANCE",
    "ASSURANCEAMERICA",
    "INFINITY AUTO INS",
    "Infinity Auto",
    "Geico Insurance",
    "OCEAN HARBOR",
    "Ocean Harbor"
    "Gateway Insurance",
    "DIRECT GENERAL",
    "NATIONAL GENERAL INSURANCE",
    "NATIONAL AUTO",
    "DIRECT AUTO",
    "Ocean Harbor Casualty Insurance Company",
    "Phildelphia Indemnity Insurance Co.",
    "STATE FARM AUTO",
    "GEICO (HOLDINGS)",
    "STATE FARM AUTO (HOLDINGS)",
    "STATE FARM AUTO (HOLDINGS CLAIM)",
    "PROGRESSIVE (HOLDINGS CLAIMS)",
    "PROGRESSIVE (HOLDINGS)"
    "Esurance",
    "National Auto Insurance",
    "Dairyland",
    "United Auto",
    "Equity",
    "Old American Insurance Company"
]

# --- Configuration ---
EXCEL_FILE_PATH = 'JanMay2019Final.XLS' # The name of your Excel file
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
PATIENT_ADDRESS_COLUMN = 'Pat Address'
PAYMENT_COLUMN = 'Payments' #
ADJUSTMENT_COLUMN = 'Adjustments' 
CREDITS_COLUMN = 'Credits'
FACILITY_COLUMN = 'Facility'
PROVIDER_COLUMN = 'Provider Name' 
CPT_UNITS_COLUMN = 'Units' # The column header for the CPT units
# --- Main Processing ---
try:
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(EXCEL_FILE_PATH)

    # --- Data Cleaning and Type Conversion ---
    # Convert date column
    df[VISIT_DATE_COLUMN] = pd.to_datetime(df[VISIT_DATE_COLUMN], errors='coerce')

    # Convert potential numeric columns, coercing errors and filling NaNs
    numeric_cols = [TOTAL_CHARGE_COLUMN, PAYMENT_COLUMN, ADJUSTMENT_COLUMN, CREDITS_COLUMN]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            print(f"Warning: Column '{col}' not found, defaulting to 0.")
            df[col] = 0 # Add column with 0 if missing

    # Ensure other relevant columns are strings
    string_cols = [ INSURANCE_COLUMN_NAME, INSURED_ID_COLUMN,
                   FIRST_NAME_COLUMN, LAST_NAME_COLUMN, CPT_COLUMN,
                   PROVIDER_COLUMN, FACILITY_COLUMN]
    for col in string_cols:
         if col in df.columns:
            df[col] = df[col].astype(str).fillna('') # Fill potential NaN in string columns
         else:
            print(f"Warning: Column '{col}' not found, using empty strings.")
            df[col] = '' # Add column with empty string if missing


    # --- Calculate Line Balance ---
    # Calculate the balance for each individual charge line
    df['LineBalance'] = df[TOTAL_CHARGE_COLUMN] - df[PAYMENT_COLUMN] - df[ADJUSTMENT_COLUMN] - df[CREDITS_COLUMN]

    # --- Filtering ---
    # Filter for car insurance rows with valid dates
    carinsurancepattern = '|'.join(AUTO_INSURANCES) # Create a regex pattern from the list
    car_insurance_df = df[
        df[INSURANCE_COLUMN_NAME].str.contains(AUTO_INSURANCES, case=False, na=False) &
        df[VISIT_DATE_COLUMN].notna() # Only include rows with valid dates
    ].copy() # Use .copy() to avoid SettingWithCopyWarning
    # Check if any car insurance rows were found
    if car_insurance_df.empty:
        print(f"No valid rows found with '{CAR_INSURANCE_KEYWORD}' in the '{INSURANCE_COLUMN_NAME}' column and valid dates.")
    else:
        # --- Grouping and Aggregation ---

        # Define the columns to include in the detailed charge list
        charge_detail_cols = [
            VISIT_DATE_COLUMN,
            CPT_COLUMN,
            PROVIDER_COLUMN,
            FACILITY_COLUMN,
            TOTAL_CHARGE_COLUMN,
            PAYMENT_COLUMN,
            ADJUSTMENT_COLUMN,
            CREDITS_COLUMN,
            CPT_UNITS_COLUMN,
            'LineBalance' # Include the calculated balance
        ]

        # Create the list of charge details for each patient
        # This converts the relevant columns of each patient's group into a list of dictionaries
        charge_details = car_insurance_df.groupby(PATIENT_ID_COLUMN).apply(
            lambda x: x[charge_detail_cols].to_dict('records'),
            include_groups=False # Avoid adding group keys to the output of apply
        ).rename('ChargeDetails') # Rename the resulting Series
               # 1. Aggregate Summary Patient Information
        patient_summary_agg = car_insurance_df.groupby(PATIENT_ID_COLUMN).agg(
            FirstName=(FIRST_NAME_COLUMN, 'first'),
            LastName=(LAST_NAME_COLUMN, 'first'),
            InsuranceProvider=(INSURANCE_COLUMN_NAME, 'first'),
            Facility=(FACILITY_COLUMN, 'first'),
            InsuranceID=(INSURED_ID_COLUMN, 'first'),
            Birthday=(PATIENT_BDAY_COLUMN, 'first'),
            FirstServiceDate=(VISIT_DATE_COLUMN, 'min'),
            LastServiceDate=(VISIT_DATE_COLUMN, 'max'),
            TotalBalance=( 'LineBalance', 'sum'), # Sum of individual line balances
            Address=(PATIENT_ADDRESS_COLUMN, 'first'),
        )
        # Combine patient info with the charge details list

        final_summary = patient_summary_agg.join(charge_details)
        # Reset index to make Patient ID a column again
        final_summary = final_summary.reset_index()


        # --- Output Results ---
        print("Summary of Car Insurance Charges per Patient (with Detailed Lines):")
        # Print the DataFrame row by row for better readability of the list column
        for index, row in final_summary.iterrows():
            print("-" * 50)
            generate_pdf(row, f"{row['FirstName']}_{row['LastName']}.pdf") # Generate PDF for each patient
            print('PDF generated successfully for patient:', row[PATIENT_ID_COLUMN])
            print(f"Patient ID: {row[PATIENT_ID_COLUMN]}")
            print(f"Name: {row['FirstName']} {row['LastName']}")
            print(f"Birthday: {row['Birthday']}")
            print(f"Insurance Provider: {row['InsuranceProvider']}")
            print(f"First Service Date: {row['FirstServiceDate']}")
            print(f"Last Service Date: {row['LastServiceDate']}")
            print(f"Total Balance: {row['TotalBalance']:.2f}")
            print(f"Facility: {row['Facility']}")
            print(f"Insurance Claim ID: {row['InsuranceID']}")
            print("Charge Details:")
            if isinstance(row['ChargeDetails'], list):
                for charge in row['ChargeDetails']:
                    # Format date within the loop for printing
                    charge_date_str = pd.to_datetime(charge[VISIT_DATE_COLUMN]).strftime('%m/%d/%Y') if pd.notna(charge[VISIT_DATE_COLUMN]) else 'N/A'
                    print(f"  - Date: {charge_date_str}, CPT: {charge[CPT_COLUMN]}, "
                          f"Provider: {charge[PROVIDER_COLUMN]}, Facility: {charge[FACILITY_COLUMN]}, "
                          f"Total: {charge[TOTAL_CHARGE_COLUMN]:.2f}, Payment: {charge[PAYMENT_COLUMN]:.2f}, "
                          f"Adjust: {charge[ADJUSTMENT_COLUMN]:.2f}, Credits: {charge[CREDITS_COLUMN]:.2f}, "
                          f"Balance: {charge['LineBalance']:.2f}")
            else:
                print("  No charge details found.")
            print("-" * 50)
        


except FileNotFoundError:
    print(f"Error: The file '{EXCEL_FILE_PATH}' was not found.")
    print("Please make sure the file name and path are correct")
except KeyError as e:
    print(f"Error: A required column was not found in the Excel file: {e}")
    print("Please ensure all specified column names exist in your file.")
    # Add all required columns to the error message for clarity
    required_cols_list = [PATIENT_ID_COLUMN, VISIT_DATE_COLUMN, CPT_COLUMN, PROVIDER_COLUMN,
                          FACILITY_COLUMN, TOTAL_CHARGE_COLUMN, PAYMENT_COLUMN, ADJUSTMENT_COLUMN,
                          CREDITS_COLUMN, INSURANCE_COLUMN_NAME, INSURED_ID_COLUMN,
                          FIRST_NAME_COLUMN, LAST_NAME_COLUMN]
    print(f"Required columns: {', '.join(required_cols_list)}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    import traceback
    traceback.print_exc() # Print detailed traceback for debugging
