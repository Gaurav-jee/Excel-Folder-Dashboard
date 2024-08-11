import pandas as pd

# Function to read the Excel table from the second sheet and extract columns
def get_table_from_excel(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name=1, engine='openpyxl')
        # Ensure there's only one row of data
        if df.shape[0] != 1:
            raise ValueError("The table should contain exactly one row of data.")
        # Convert the row to a dictionary
        row = df.fillna(0).replace("-", 0).iloc[0].to_dict()
        return row
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return {}
