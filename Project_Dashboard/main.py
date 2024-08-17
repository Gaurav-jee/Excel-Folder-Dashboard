import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import platform
from datetime import datetime
from file_utils import get_file_details, write_cache, read_cache, get_table_from_excel
from ui_utils import (setup_modern_style, setup_treeview, add_scrollbars, create_rounded_button,
                      create_export_frame, create_percentage_label)
from json_utils import get_default_paths, update_default_paths

# Initialize the main application window
root = tk.Tk()
root.title("Excel Folder Dashboard")
root.geometry("1000x700")
setup_modern_style()

# Load default paths
default_folder_path, default_cache_path = get_default_paths()

# Variables to store paths
folder_path = tk.StringVar(value=default_folder_path)
cache_path = tk.StringVar(value=default_cache_path)

# Define initial columns and additional columns from the Excel table
predefined_columns = ["S.L.", "Name of the file", "Date of Creation", "Date of Modification", "File Path"]
additional_columns = [
    "TOTAL NO PCS.", "TOTAL AMOUNT", "मूर्ति दुकान", "गणेश लक्ष्मी दुकान",
    "LOADING CHARGE", "TRANSPORTATION", "PACKING CHARGES", "पहेले का बकाया",
    "GRAND TOTAL", "ADVANCE", "PAYABLE"
]

# Combine predefined columns with additional columns
columns = predefined_columns + additional_columns

# Function to browse for folder
def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path.set(folder_selected)
        update_default_paths(folder_selected, cache_path.get())

# Function to browse for cache file
def browse_cache():
    cache_file = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if cache_file:
        cache_path.set(cache_file)
        update_default_paths(folder_path.get(), cache_file)

# Create and pack UI elements for path selection
path_frame = ttk.Frame(root)
path_frame.pack(pady=10, padx=10, fill=tk.X)

ttk.Label(path_frame, text="Folder Path:").grid(row=0, column=0, sticky="w", pady=5)
ttk.Entry(path_frame, textvariable=folder_path, width=50).grid(row=0, column=1, padx=5)
create_rounded_button(path_frame, "Browse", browse_folder).grid(row=0, column=2)

ttk.Label(path_frame, text="Cache Path:").grid(row=1, column=0, sticky="w", pady=5)
ttk.Entry(path_frame, textvariable=cache_path, width=50).grid(row=1, column=1, padx=5)
create_rounded_button(path_frame, "Browse", browse_cache).grid(row=1, column=2)

# Create a Frame for the Treeview and Scrollbars
tree_frame = ttk.Frame(root)
tree_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

# Setup Treeview
tree = setup_treeview(tree_frame, columns)
add_scrollbars(tree_frame, tree)

# Create export frame
create_export_frame(root, tree, columns)

# Create progress bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, style="TProgressbar")
progress_bar.pack(fill=tk.X, padx=10, pady=5)

# Create the percentage label
percentage_label = create_percentage_label(root)

def update_progress(current, total):
    progress = (current / total) * 100
    progress_var.set(progress)
    percentage_label.config(text=f"{int(progress)}%")
    root.update_idletasks()

def get_file_details_with_progress(folder_path):
    file_list = []
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    total_files = len(files)
    for index, filename in enumerate(files):
        filepath = os.path.join(folder_path, filename)
        if os.path.isfile(filepath):
            creation_time = os.path.getctime(filepath)
            modification_time = os.path.getmtime(filepath)
            table_data, success = get_table_from_excel(filepath)
            if success:
                file_list.append({
                    'S.L.': index + 1,
                    'Name of the file': filename,
                    'Date of Creation': datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S'),
                    'Date of Modification': datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S'),
                    **table_data,
                    'File Path': filepath
                })
        update_progress(index + 1, total_files)
    return file_list

# Function to update the Treeview with new data
def update_treeview(data):
    for i in tree.get_children():
        tree.delete(i)
    for file in data:
        row_data = [file.get(col, '') for col in columns]
        tree.insert("", tk.END, values=tuple(row_data), tags=(file['File Path'],))

# Function to refresh the data
def refresh_data():
    if not folder_path.get() or not cache_path.get():
        messagebox.showerror("Error", "Please select both folder and cache paths.")
        return
    file_data = get_file_details_with_progress(folder_path.get())
    update_treeview(file_data)
    write_cache(file_data, cache_path.get())
    messagebox.showinfo("Success", "Data refreshed and cache updated.")

# Create the refresh button
refresh_button = create_rounded_button(root, "Refresh", refresh_data)
refresh_button.pack(pady=10)

# Function to remove rows with "new" in the filename
def remove_new_files():
    new_keywords = ["new", "न्यू"]
    items_to_remove = []
    for item in tree.get_children():
        values = tree.item(item)['values']
        filename = values[columns.index("Name of the file")].lower()
        if any(keyword in filename for keyword in new_keywords):
            items_to_remove.append(item)
    
    for item in items_to_remove:
        tree.delete(item)
    
    messagebox.showinfo("Info", f"Removed {len(items_to_remove)} files containing 'new' or 'न्यू' in the filename.")

# Create the "Remove New Files" button
remove_new_button = create_rounded_button(root, "Remove New Files", remove_new_files)
remove_new_button.pack(pady=10)

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

# Read cached data or load new data if cache is empty
def load_initial_data():
    if cache_path.get():
        cached_data = read_cache(cache_path.get())
        if cached_data:
            update_treeview(cached_data)
        else:
            messagebox.showinfo("Info", "No cached data found. Please refresh to load data.")
    else:
        messagebox.showinfo("Info", "Please select a cache file and refresh to load data.")

# Load initial data button
load_button = create_rounded_button(root, "Load Cached Data", load_initial_data)
load_button.pack(pady=10)

# Run the application
root.mainloop()