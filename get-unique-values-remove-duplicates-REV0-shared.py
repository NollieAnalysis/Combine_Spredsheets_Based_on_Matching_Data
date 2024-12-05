# import
import pandas as pd

# user input spreadsheet
user_first_spread_name = str(input("Enter name of prep spreadsheet file (including file type): "))

# path to spreadsheet
first_spreadsheet_path = f"/{user_first_spread_name}"

# output spreadsheet path
output_spreadsheet_path = f"/{user_first_spread_name}(ADDED_INFO).xlsx"

# key column to remove duplicates of
key_column_first = 'Order Number'

# custom function
def get_unique_values(first_spreadsheet_path, key_column_first, output_spreadsheet_path):
    df = pd.read_excel(first_spreadsheet_path)
    df1 = df.drop_duplicates(subset=key_column_first)
    df1.to_excel(output_spreadsheet_path, index=False)
    return df1

# call custom function
get_unique_values(first_spreadsheet_path, key_column_first, output_spreadsheet_path)