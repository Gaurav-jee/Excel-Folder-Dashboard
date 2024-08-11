# 'C:/ART_GALLERYHPWHITE/2023/NEW BILLS'

import os
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import subprocess
import platform

# Function to get file details
def get_file_details(folder_path):
    file_list = []
    for index, filename in enumerate(os.listdir(folder_path)):
        filepath = os.path.join(folder_path, filename)
        if os.path.isfile(filepath):
            creation_time = os.path.getctime(filepath)
            modification_time = os.path.getmtime(filepath)
            file_list.append({
                'S.L.': index + 1,
                'Name of the file': filename,
                'Date of Creation': datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S'),
                'Date of Modification': datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S'),
                'File Path': filepath
            })
    return file_list

# Folder path (change this to your target folder)
folder_path = 'C:/ART_GALLERYHPWHITE/2023/NEW BILLS'

# Create a list of file details
file_data = get_file_details(folder_path)

# Create the main application window
root = tk.Tk()
root.title("File Dashboard")

# Create a Treeview to display the file details
tree = ttk.Treeview(root, columns=("S.L.", "Name of the file", "Date of Creation", "Date of Modification"), show='headings')
tree.heading("S.L.", text="S.L.")
tree.heading("Name of the file", text="Name of the file")
tree.heading("Date of Creation", text="Date of Creation")
tree.heading("Date of Modification", text="Date of Modification")

# Insert the file details into the Treeview
for file in file_data:
    tree.insert("", tk.END, values=(file['S.L.'], file['Name of the file'], file['Date of Creation'], file['Date of Modification']), tags=(file['File Path'],))

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

# Pack the Treeview into the window
tree.pack(expand=True, fill=tk.BOTH)

# Run the application
root.mainloop()
