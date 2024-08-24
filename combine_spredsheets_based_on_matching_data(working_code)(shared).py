########################################################################
# combine spredsheets based on matching data
########################################################################


# packages
import pandas as pd


###########################
# extract data from first spreadsheet in order to add as filter in other software
###########################

def list_of_order_numbers(first_spreadsheet_path, output_spreadsheet_path, key_column_first, sheet_names=None):
    all_orders = []  # List to hold all numeric orders from all sheets
    
    for sheet in sheet_names:
        # Load the specific sheet into a DataFrame
        df = pd.read_excel(first_spreadsheet_path, sheet_name=sheet)
        
        # Attempt to convert specific column to numeric values, forcing non-numeric values to NaN
        df[''] = pd.to_numeric(df[''], errors='coerce')
        
        # Drop rows where specific column is NaN (i.e., non-numeric values)
        numeric_orders = df[''].dropna()
        
        # Append the cleaned orders to the list
        all_orders.append(numeric_orders)
    
    # Concatenate all numeric orders from all sheets into a single Series
    combined_orders = pd.concat(all_orders).reset_index(drop=True)

    # Convert the numeric orders to strings and add a comma after each value
    combined_orders = combined_orders.apply(lambda x: f"{int(x)},")

    # Save the combined orders to the output spreadsheet
    combined_orders.to_excel(output_spreadsheet_path, index=False, header=False)

    print(f"Spreadsheet saved as {output_spreadsheet_path}")



# Example usage
first_spreadsheet_path = ""
second_spreadsheet_path = ""
output_spreadsheet_path = ""
key_column_first = ''  # The column to match in the first spreadsheet
key_column_second = ''  # The column to match in the second spreadsheet
columns_to_copy = []  # Columns from second spreadsheet to copy to first
sheet_names = []  # The name or index of the sheets you want to use in the first spreadsheet


list_of_order_numbers(first_spreadsheet_path, output_spreadsheet_path, key_column_first, sheet_names)






###########################
# working code to add columns to first spreadsheet and copy information from a second spreadsheet to the first spreadsheet where columns were added
# added capability to choose specific sheet / tabs (in cases with multiple sheets / tabs) in first spreadsheet
# added capability to add three new columns to specific sheets where specific information from first and second spreadsheets match
###########################

def combine_spreadsheets(first_spreadsheet_path, second_spreadsheet_path, output_spreadsheet_path, key_column_first, key_column_second, columns_to_copy, sheet_names=None):
    # Load the second spreadsheet
    df2 = pd.read_excel(second_spreadsheet_path)

    # Dictionary to hold DataFrames for each sheet
    combined_sheets = {}

    for sheet_name in sheet_names:
        # Load the specific sheet from the first spreadsheet
        df1 = pd.read_excel(first_spreadsheet_path, sheet_name=sheet_name)

        # Use a merge operation to match and copy data
        df_combined = pd.merge(df1, df2[[key_column_second] + columns_to_copy], how='left', left_on=key_column_first, right_on=key_column_second)

        # Drop the redundant key column from df2 that was added during merge
        df_combined = df_combined.drop(columns=[key_column_second])

        # Store the combined DataFrame in the dictionary
        combined_sheets[sheet_name] = df_combined

    # Save all combined DataFrames to the output spreadsheet
    with pd.ExcelWriter(output_spreadsheet_path) as writer:
        for sheet_name, df_combined in combined_sheets.items():
            df_combined.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Combined spreadsheet saved as {output_spreadsheet_path}")




# Example usage
first_spreadsheet_path = ""
second_spreadsheet_path = ""
output_spreadsheet_path = ""
key_column_first = ''  # The column to match in the first spreadsheet
key_column_second = ''  # The column to match in the second spreadsheet
columns_to_copy = []  # Columns from second spreadsheet to copy to first
sheet_names = []  # The name or index of the sheets you want to use in the first spreadsheet

combine_spreadsheets(first_spreadsheet_path, second_spreadsheet_path, output_spreadsheet_path, key_column_first, key_column_second, columns_to_copy, sheet_names)