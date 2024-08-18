import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from excel_utils import get_table_from_excel

def get_table_from_excel(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name=1, engine='openpyxl')
        if df.shape[0] != 1:
            raise ValueError("The table should contain exactly one row of data.")
        
        # Strip whitespace from column names and convert to lowercase
        df.columns = df.columns.str.strip().str.lower()
        
        # Create a mapping of lowercase column names to new concise English names
        column_mapping = {
            'total no pcs.': 'total_pieces',
            'total amount': 'total_amount',
            'मूर्ति दुकान': 'murti_shop',
            'गणेश लक्ष्मी दुकान': 'ganesh_lakshmi_shop',
            'loading charge': 'loading_charge',
            'transportation': 'transportation',
            'packing charges': 'packing_charges',
            'पहेले का बकाया': 'previous_balance',
            'grand total': 'grand_total',
            'advance': 'advance',
            'payable': 'payable'
        }
        
        # Rename columns using the mapping
        df.rename(columns=column_mapping, inplace=True)
        
        # Define data types for each column
        data_types = {
            'total_pieces': 'int',
            'total_amount': 'float',
            'murti_shop': 'float',
            'ganesh_lakshmi_shop': 'float',
            'loading_charge': 'float',
            'transportation': 'float',
            'packing_charges': 'float',
            'previous_balance': 'float',
            'grand_total': 'float',
            'advance': 'float',
            'payable': 'float'
        }
        
        # Convert data types and replace "-" with 0
        for col, dtype in data_types.items():
            if col in df.columns:
                df[col] = df[col].replace("-", 0)
                df[col] = df[col].astype(dtype)
        
        # Convert the row to a dictionary
        row = df.iloc[0].to_dict()
        
        return row, True
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return {}, False

def get_file_details(folder_path):
    file_list = []
    for index, filename in enumerate(os.listdir(folder_path)):
        filepath = os.path.join(folder_path, filename)
        if os.path.isfile(filepath) and filename.endswith('.xlsx'):
            creation_time = os.path.getctime(filepath)
            modification_time = os.path.getmtime(filepath)
            table_data, success = get_table_from_excel(filepath)  # Unpack both values
            if success:  # Check if the data was read successfully
                file_list.append({
                    'S.L.': index + 1,
                    'Name of the file': filename,
                    'Date of Creation': datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S'),
                    'Date of Modification': datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S'),
                    'TOTAL NO PCS.': table_data.get('TOTAL NO PCS.', 0),
                    'TOTAL AMOUNT': table_data.get('TOTAL AMOUNT ', 0),
                    'मूर्ति दुकान': table_data.get('मूर्ति दुकान', 0),
                    'गणेश लक्ष्मी दुकान': table_data.get('गणेश लक्ष्मी दुकान', 0),
                    'LOADING CHARGE': table_data.get('LOADING CHARGE', 0),
                    'TRANSPORTATION': table_data.get('TRANSPORTATION', 0),
                    'PACKING CHARGES': table_data.get('PACKING CHARGES', 0),
                    'पहेले का बकाया': table_data.get('पहेले का बकाया', 0),
                    'GRAND TOTAL': table_data.get('GRAND TOTAL', 0),
                    'ADVANCE': table_data.get('ADVANCE', 0),
                    'PAYABLE': table_data.get('PAYABLE', 0),
                    'File Path': filepath
                })
    return file_list


# Write cache function
def write_cache(data, cache_path):
    try:
        # Convert data to a DataFrame
        df = pd.DataFrame(data)
        # Write to Excel with UTF-8 encoding
        df.to_excel(cache_path, index=False, engine='openpyxl')
    except Exception as e:
        print(f"Error writing to cache: {e}")



def read_cache(cache_path):
    try:
        # Read from Excel with UTF-8 encoding
        df = pd.read_excel(cache_path)
        return df.to_dict('records')
    except Exception as e:
        print(f"Error reading from cache: {e}")
        return []

