import os
import pandas as pd
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import subprocess
import platform

# Function to get file details and the Excel table
def get_file_details(folder_path):
    file_list = []
    for index, filename in enumerate(os.listdir(folder_path)):
        filepath = os.path.join(folder_path, filename)
        if os.path.isfile(filepath) and filename.endswith('.xlsx'):
            creation_time = os.path.getctime(filepath)
            modification_time = os.path.getmtime(filepath)
            table_data = get_table_from_excel(filepath)
            file_list.append({
                'S.L.': index + 1,
                'Name of the file': filename,
                'Date of Creation': datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S'),
                'Date of Modification': datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S'),
                'TOTAL NO PCS.': table_data.get('TOTAL NO PCS.', 0),
                'TOTAL AMOUNT': table_data.get('TOTAL AMOUNT', 0),
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

# Folder path (change this to your target folder)
folder_path = 'C:/ART_GALLERYHPWHITE/2023/NEW BILLS'

# Create a list of file details
file_data = get_file_details(folder_path)

# Initialize the main application window
root = tk.Tk()
root.title("File Dashboard")

# Define initial columns and additional columns from the Excel table
predefined_columns = ["S.L.", "Name of the file", "Date of Creation", "Date of Modification", "File Path"]
additional_columns = [
    "TOTAL NO PCS.", "TOTAL AMOUNT", "मूर्ति दुकान", "गणेश लक्ष्मी दुकान",
    "LOADING CHARGE", "TRANSPORTATION", "PACKING CHARGES", "पहेले का बकाया",
    "GRAND TOTAL", "ADVANCE", "PAYABLE"
]

# Combine predefined columns with additional columns
columns = predefined_columns + additional_columns

# Create a Frame for the Treeview and Scrollbars
frame = ttk.Frame(root)
frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

# Create a Treeview with dynamic columns
tree = ttk.Treeview(frame, columns=columns, show='headings')

# Add a border to the Treeview
tree.config(style="Custom.Treeview")

# Set the column headings
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150)  # Adjust the width as needed

# Insert the file details into the Treeview
for file in file_data:
    # Prepare row data with both predefined and additional columns
    row_data = [file.get(col, '') for col in columns]
    
    # Insert the row into the Treeview
    tree.insert("", tk.END, values=tuple(row_data), tags=(file['File Path'],))

# Bind double-click event to open the file
def open_file(event):
    item = tree.selection()[0]
    file_path = tree.item(item, "tags")[0]
    if platform.system() == "Windows":
        os.startfile(file_path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", file_path], check=True)
    else:  # Linux
        subprocess.run(["xdg-open", file_path], check=True)

tree.bind("<Double-1>", open_file)

# Create vertical and horizontal scrollbars
vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

# Pack the scrollbars and Treeview
vsb.pack(side="right", fill="y")
hsb.pack(side="bottom", fill="x")
tree.pack(expand=True, fill=tk.BOTH)

# Style the Treeview to make headers pop and apply a modern look
style = ttk.Style()
style.configure("Custom.Treeview.Heading",
                font=("Helvetica", 12, "bold"),
                background="light blue",
                foreground="white",
                relief="raised")
style.map("Custom.Treeview.Heading",
          background=[('active', 'neon green'), ('hover', 'neon green')],
          foreground=[('active', 'white'), ('hover', 'white')])
style.configure("Custom.Treeview", highlightthickness=1, bd=1, relief="solid")

# Run the application
root.mainloop()
