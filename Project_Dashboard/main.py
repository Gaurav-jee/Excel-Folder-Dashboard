import os
import tkinter as tk
from tkinter import ttk
import subprocess
import platform
from datetime import datetime
from file_utils import get_file_details, write_cache, read_cache, get_table_from_excel
from ui_utils import setup_treeview, add_scrollbars_and_style, create_percentage_label

# Folder path (change this to your target folder)
folder_path = 'C:/ART_GALLERYHPWHITE/2023/NEW BILLS'

# Define the cache path
cache_path = 'C:/ART_GALLERYHPWHITE/Experiments/Folder Dashboard/Project_Dashboard/cache.xlsx'

# Define initial columns and additional columns from the Excel table
predefined_columns = ["S.L.", "Name of the file", "Date of Creation", "Date of Modification", "File Path"]
additional_columns = [
    "TOTAL NO PCS.", "TOTAL AMOUNT", "मूर्ति दुकान", "गणेश लक्ष्मी दुकान",
    "LOADING CHARGE", "TRANSPORTATION", "PACKING CHARGES", "पहेले का बकाया",
    "GRAND TOTAL", "ADVANCE", "PAYABLE"
]

# Combine predefined columns with additional columns
columns = predefined_columns + additional_columns

# Function to refresh the data with percentage update
def refresh_data():
    root.after(100, lambda: update_data(percentage_label))

def update_data(percentage_label):
    # Get file details
    file_data = get_file_details(folder_path)
    
    # Display data and update cache
    percentage_label.config(text="100%")
    
    # Write to cache
    write_cache(file_data, cache_path)


# Function to get file details and update percentage
def get_file_details_with_percentage(folder_path, percentage_label):
    file_list = []
    files = os.listdir(folder_path)
    total_files = len(files)
    for index, filename in enumerate(files):
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
                **table_data,
                'File Path': filepath
            })
        percentage = (index + 1) / total_files * 100
        percentage_label.config(text=f"{int(percentage)}%")
        root.update_idletasks()  # Update the UI
    return file_list

# Function to update the Treeview with new data
def update_treeview(data):
    for i in tree.get_children():
        tree.delete(i)
    for file in data:
        row_data = [file.get(col, '') for col in columns]
        tree.insert("", tk.END, values=tuple(row_data), tags=(file['File Path'],))

# Initialize the main application window
root = tk.Tk()
root.title("File Dashboard")

# Create a Frame for the Treeview and Scrollbars
frame = ttk.Frame(root)
frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

# Setup Treeview
tree = setup_treeview(frame, columns)
add_scrollbars_and_style(frame, tree)

# Create a style for the refresh button
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12), padding=6)
style.configure("TButton", background="light blue")  # Default background color
style.map("TButton",
          background=[("active", "neon green"), ("!active", "light blue")])

# Create the refresh button
refresh_button = ttk.Button(root, text="Refresh", command=refresh_data, style="TButton")
refresh_button.pack(pady=10)

# Hover effects for the refresh button using direct color change
def on_enter(event):
    event.widget.state(["!pressed", "active"])

def on_leave(event):
    event.widget.state(["!pressed", "!active"])

refresh_button.bind("<Enter>", on_enter)
refresh_button.bind("<Leave>", on_leave)

# Create the percentage label
percentage_label = create_percentage_label(root)

# Read cached data or load new data if cache is empty
cached_data = read_cache(cache_path)
if cached_data:
    update_treeview(cached_data)
else:
    refresh_data()

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

# Run the application
root.mainloop()
