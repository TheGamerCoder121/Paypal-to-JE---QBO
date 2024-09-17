# import pandas as pd
# import os
# from datetime import datetime

# # Get the current directory
# current_dir = os.path.dirname(os.path.abspath(__file__))

# # Construct the full path to the Excel file
# excel_file_path = os.path.join(current_dir, 'attachments', 'KKS-PayPal+August.xlsx')

# # Load the Excel file
# df = pd.read_excel(excel_file_path)

# # Display the first few rows to verify the data
# print(df.head())

# # Display column names
# print("\nColumns:")
# print(df.columns)

# # Display basic information about the DataFrame
# print("\nDataFrame Info:")
# df.info()

# # Data processing and transformation logic
# def categorize_transaction(row):
#     description = row['Description']
#     amount = row['Net']

#     if 'Payment' in description or 'Deposit' in description:
#         if amount > 0:
#             return 'Income'
#         else:
#             return 'Expense'
#     elif 'Withdrawal' in description:
#         return 'Expense'
#     elif 'Refund' in description:
#         return 'Income'
#     else:
#         return 'Other'

# def get_account(category):
#     if category == 'Income':
#         return 'PayPal Income'
#     elif category == 'Expense':
#         return 'PayPal Expense'
#     else:
#         return 'PayPal Other'

# # Group transactions by date and absolute amount
# def group_transactions(df):
#     # Convert Date to datetime.date and create Abs_Amount column
#     df['Date'] = pd.to_datetime(df['Date']).dt.date
#     df['Abs_Amount'] = df['Net'].abs()
#     # Group by Date and Abs_Amount to match transactions with similar dates and exact absolute amounts
#     return df.groupby(['Date', 'Abs_Amount'])

# # Create journal entries from grouped transactions
# def create_journal_entries(grouped_transactions):
#     journal_entries = []
#     journal_no = 1

#     for (date, abs_amount), group in grouped_transactions:
#         # Skip groups with net zero transactions
#         if abs(group['Net'].sum()) < 0.01:  # Allow for small rounding errors
#             continue

#         num_transactions = len(group)
#         net_amount = group['Net'].sum()

#         # Create main entry for PayPal Balance
#         main_entry = {
#             'Journal No': journal_no,
#             'Journal Date': date.strftime('%m/%d/%y'),
#             'Memo': f"PayPal - {num_transactions} Transaction{'s' if num_transactions > 1 else ''}",
#             'Account': 'PayPal Balance',
#             'Amount': net_amount,
#             'Description': ', '.join(group['Description']),
#             'Name': ', '.join(group['Name'].dropna()),
#             'Location': '',
#             'Class': '',
#             'Currency Code': group['Currency'].iloc[0],
#             'Exchange Rate': '',
#             'Is Adjustment': 'FALSE'
#         }
#         journal_entries.append(main_entry)

#         # Create offsetting entries
#         for _, row in group.iterrows():
#             category = categorize_transaction(row)
#             account = get_account(category)

#             offsetting_entry = main_entry.copy()
#             offsetting_entry['Account'] = account
#             offsetting_entry['Amount'] = -row['Net']
#             offsetting_entry['Description'] = row['Description']
#             offsetting_entry['Name'] = row['Name'] if pd.notna(row['Name']) else ''
#             journal_entries.append(offsetting_entry)

#         journal_no += 1

#     return pd.DataFrame(journal_entries)

# # Group transactions and create journal entries
# grouped_transactions = group_transactions(df)
# journal_df = create_journal_entries(grouped_transactions)

# # Save the new DataFrame to an Excel file
# output_excel_path = os.path.join(current_dir, 'PayPal_Journal_Entries.xlsx')
# journal_df.to_excel(output_excel_path, index=False)

# print(f"\nJournal entries have been saved to: {output_excel_path}")
# print("\nFirst few rows of the journal entries:")
# print(journal_df.head())

# # Convert Excel to CSV
# output_csv_path = os.path.join(current_dir, 'PayPal_Journal_Entries.csv')
# journal_df.to_csv(output_csv_path, index=False)

# print(f"\nJournal entries have been saved as CSV to: {output_csv_path}")
# print("\nFirst few rows of the CSV file:")
# print(pd.read_csv(output_csv_path).head())

# def verify_data_transformation():
#     csv_df = pd.read_csv(output_csv_path)

#     # Check if all required columns are present
#     required_columns = ['Journal No', 'Journal Date', 'Memo', 'Account', 'Amount', 'Description', 'Name', 'Location', 'Class', 'Currency Code', 'Exchange Rate', 'Is Adjustment']
#     missing_columns = set(required_columns) - set(csv_df.columns)
#     if missing_columns:
#         print(f"Warning: Missing columns in the CSV file: {missing_columns}")
#     else:
#         print("All required columns are present in the CSV file.")

#     # Check for non-empty values in important fields
#     important_fields = ['Journal No', 'Journal Date', 'Account', 'Amount', 'Description']
#     for field in important_fields:
#         empty_count = csv_df[field].isna().sum()
#         if empty_count > 0:
#             print(f"Warning: {empty_count} empty values found in the '{field}' column.")

#     # Check if amounts balance out for each journal entry
#     journal_balance = csv_df.groupby('Journal No')['Amount'].sum()
#     unbalanced_entries = journal_balance[abs(journal_balance) > 0.01]  # Allow for small rounding errors
#     if not unbalanced_entries.empty:
#         print(f"Warning: Unbalanced journal entries found: {unbalanced_entries}")
#     else:
#         print("All journal entries are balanced.")

#     print("\nData verification complete.")

# verify_data_transformation()
import pandas as pd
import os
from datetime import datetime

def create_input_directory():
    input_dir = 'input_files'
    os.makedirs(input_dir, exist_ok=True)
    return input_dir

def list_excel_files(directory):
    return [f for f in os.listdir(directory) if f.endswith('.xlsx')]

def select_input_file(files):
    print("Available Excel files:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")

    while True:
        try:
            choice = int(input("Enter the number of the file you want to process: "))
            if 1 <= choice <= len(files):
                return files[choice - 1]
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def get_input_file():
    input_dir = create_input_directory()
    excel_files = list_excel_files(input_dir)

    if not excel_files:
        print(f"No Excel files found in the '{input_dir}' directory. Please add some files and try again.")
        exit(1)

    selected_file = select_input_file(excel_files)
    return os.path.join(input_dir, selected_file)

def create_output_directory():
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

# Get input file and create output directory
excel_file_path = get_input_file()
output_folder = create_output_directory()

# Load the selected Excel file
df = pd.read_excel(excel_file_path)

# Display the first few rows to verify the data
print(df.head())

# Display column names
print("\nColumns:")
print(df.columns)

# Display basic information about the DataFrame
print("\nDataFrame Info:")
df.info()

# Data processing and transformation logic
def categorize_transaction(row):
    description = row['Description']
    amount = row['Net']

    if 'Payment' in description or 'Deposit' in description:
        if amount > 0:
            return 'Income'
        else:
            return 'Expense'
    elif 'Withdrawal' in description:
        return 'Expense'
    elif 'Refund' in description:
        return 'Income'
    else:
        return 'Other'

def get_account(category):
    if category == 'Income':
        return 'PayPal Income'
    elif category == 'Expense':
        return 'PayPal Expense'
    else:
        return 'PayPal Other'

# Group transactions by date and absolute amount
def group_transactions(df):
    # Convert Date to datetime.date and create Abs_Amount column
    df['Date'] = pd.to_datetime(df['Date']).dt.date
    df['Abs_Amount'] = df['Net'].abs()
    # Group by Date and Abs_Amount to match transactions with similar dates and exact absolute amounts
    return df.groupby(['Date', 'Abs_Amount'])

# Create journal entries from grouped transactions
def create_journal_entries(grouped_transactions):
    journal_entries = []
    journal_no = 1
    flagged_entries = []

    def identify_source(description):
        if '1724' in description:
            return '1724 card'
        elif '3009' in description:
            return '3009 account'
        elif '3001' in description:
            return '3001 account'
        else:
            return 'PayPal Balance'

    for (date, abs_amount), group in grouped_transactions:
        # Skip groups with net zero transactions
        if abs(group['Net'].sum()) < 0.01:  # Allow for small rounding errors
            continue

        num_transactions = len(group)
        net_amount = group['Net'].sum()

        # Identify the sources of money
        sources = group['Description'].apply(identify_source)
        main_source = sources.mode().iloc[0]

        # Check for partial payments from PayPal Balance
        partial_payment = 'PayPal Balance' in sources.values and len(sources.unique()) > 1

        # Create main entry
        main_entry = {
            'Journal No': journal_no,
            'Journal Date': date.strftime('%m/%d/%y'),
            'Memo': f"PayPal - {num_transactions} Transaction{'s' if num_transactions > 1 else ''}",
            'Account': main_source,
            'Amount': net_amount,
            'Description': ', '.join(group['Description']),
            'Name': ', '.join(group['Name'].dropna()),
            'Location': '',
            'Class': '',
            'Currency Code': group['Currency'].iloc[0],
            'Exchange Rate': '',
            'Is Adjustment': 'TRUE' if partial_payment else 'FALSE'
        }
        journal_entries.append(main_entry)

        # Create offsetting entries
        offsetting_amount = 0
        for idx, row in group.iterrows():
            category = categorize_transaction(row)
            account = get_account(category)

            offsetting_entry = main_entry.copy()
            offsetting_entry['Journal No'] = '' if idx > 0 else journal_no
            offsetting_entry['Journal Date'] = '' if idx > 0 else main_entry['Journal Date']
            offsetting_entry['Memo'] = ''
            offsetting_entry['Account'] = account
            offsetting_entry['Amount'] = -row['Net']
            offsetting_entry['Description'] = row['Description']
            offsetting_entry['Name'] = row['Name'] if pd.notna(row['Name']) else ''
            offsetting_entry['Is Adjustment'] = ''
            journal_entries.append(offsetting_entry)

            offsetting_amount += -row['Net']

        # Check if entries are balanced
        if abs(net_amount + offsetting_amount) > 0.01 or partial_payment:  # Allow for small rounding errors
            flagged_entries.append({
                'Journal No': journal_no,
                'Date': date,
                'Net Amount': net_amount,
                'Offsetting Amount': offsetting_amount,
                'Difference': net_amount + offsetting_amount,
                'Partial Payment': 'Yes' if partial_payment else 'No'
            })

        journal_no += 1

    # Print flagged entries for manual review
    if flagged_entries:
        print("\nFlagged entries for manual review:")
        for entry in flagged_entries:
            print(f"Journal No: {entry['Journal No']}, Date: {entry['Date']}, "
                  f"Difference: {entry['Difference']:.2f}, Partial Payment: {entry['Partial Payment']}")

    return pd.DataFrame(journal_entries)

# Group transactions and create journal entries
grouped_transactions = group_transactions(df)
journal_df = create_journal_entries(grouped_transactions)

# Save the new DataFrame to an Excel file
output_excel_path = os.path.join(output_folder, 'PayPal_Journal_Entries.xlsx')
journal_df.to_excel(output_excel_path, index=False)

print(f"\nJournal entries have been saved to: {output_excel_path}")
print("\nFirst few rows of the journal entries:")
print(journal_df.head())

# Convert Excel to CSV
output_csv_path = os.path.join(output_folder, 'PayPal_Journal_Entries.csv')
journal_df.to_csv(output_csv_path, index=False)

print(f"\nJournal entries have been saved as CSV to: {output_csv_path}")
print("\nFirst few rows of the CSV file:")
print(pd.read_csv(output_csv_path).head())

def verify_data_transformation():
    csv_df = pd.read_csv(output_csv_path)

    # Check if all required columns are present
    required_columns = ['Journal No', 'Journal Date', 'Memo', 'Account', 'Amount', 'Description', 'Name', 'Location', 'Class', 'Currency Code', 'Exchange Rate', 'Is Adjustment']
    missing_columns = set(required_columns) - set(csv_df.columns)
    if missing_columns:
        print(f"Warning: Missing columns in the CSV file: {missing_columns}")
    else:
        print("All required columns are present in the CSV file.")

    # Check for non-empty values in important fields
    important_fields = ['Journal No', 'Journal Date', 'Account', 'Amount', 'Description']
    for field in important_fields:
        empty_count = csv_df[field].isna().sum()
        if empty_count > 0:
            print(f"Warning: {empty_count} empty values found in the '{field}' column.")

    # Check if amounts balance out for each journal entry
    journal_balance = csv_df.groupby('Journal No')['Amount'].sum()
    unbalanced_entries = journal_balance[abs(journal_balance) > 0.01]  # Allow for small rounding errors
    if not unbalanced_entries.empty:
        print(f"Warning: Unbalanced journal entries found: {unbalanced_entries}")
    else:
        print("All journal entries are balanced.")

    print("\nData verification complete.")

verify_data_transformation()
