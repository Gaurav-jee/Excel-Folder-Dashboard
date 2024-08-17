import tkinter as tk
from tkinter import ttk, filedialog
import csv
import json

def setup_modern_style():
    style = ttk.Style()
    style.theme_use('clam')

    style.configure("Treeview",
                    background="#F0F0F0",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="#F0F0F0",
                    bordercolor="#E0E0E0",
                    borderwidth=0)
    style.map('Treeview', background=[('selected', '#3498db')])

    style.configure("Treeview.Heading",
                    background="#3498db",
                    foreground="white",
                    relief="flat")
    style.map("Treeview.Heading",
              background=[('active', '#2980b9')])

    # ... (rest of the style configurations remain the same)

def setup_treeview(frame, columns):
    tree = ttk.Treeview(frame, columns=columns, show='headings')
    for col in columns:
        tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))
        tree.column(col, width=150, anchor="center")
    return tree

def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    try:
        l.sort(key=lambda t: float(t[0]), reverse=reverse)
    except ValueError:
        l.sort(key=lambda t: t[0].lower(), reverse=reverse)

    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def add_scrollbars(frame, tree):
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(expand=True, fill=tk.BOTH)

def create_rounded_button(parent, text, command):
    button = tk.Button(parent, text=text, command=command, relief="flat", bg="#3498db", fg="white",
                       activebackground="#2980b9", activeforeground="white", bd=0, padx=10, pady=5,
                       font=("Helvetica", 10, "bold"))
    button.bind("<Enter>", lambda e: e.widget.config(bg="#2980b9"))
    button.bind("<Leave>", lambda e: e.widget.config(bg="#3498db"))
    return button

def create_export_frame(parent, tree, columns):
    export_frame = ttk.Frame(parent)
    export_frame.pack(pady=10, padx=10, fill=tk.X)

    def export_to_csv():
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(columns)
                for item in tree.get_children():
                    writer.writerow(tree.item(item)['values'])

    def export_to_json():
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            data = []
            for item in tree.get_children():
                values = tree.item(item)['values']
                data.append(dict(zip(columns, values)))
            with open(file_path, 'w', encoding='utf-8') as jsonfile:
                json.dump(data, jsonfile, indent=2, ensure_ascii=False)

    csv_button = create_rounded_button(export_frame, "Export to CSV", export_to_csv)
    csv_button.pack(side=tk.LEFT, padx=(0, 5))

    json_button = create_rounded_button(export_frame, "Export to JSON", export_to_json)
    json_button.pack(side=tk.LEFT, padx=(0, 5))

def create_percentage_label(root):
    percentage_label = ttk.Label(root, text="0%", font=("Helvetica", 10, "bold"))
    percentage_label.pack(side=tk.TOP, padx=10, pady=5)
    return percentage_label