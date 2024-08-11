import tkinter as tk
from tkinter import ttk

from tkinter import ttk

def setup_treeview(frame, columns):
    tree = ttk.Treeview(frame, columns=columns, show='headings')
    tree.config(style="Custom.Treeview")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)
    return tree

def add_scrollbars_and_style(frame, tree):
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(expand=True, fill=tk.BOTH)
    
    style = ttk.Style()
    style.configure("Custom.Treeview.Heading",
                    font=("Helvetica", 12, "bold"),
                    background="light blue",
                    foreground="black",
                    relief="raised")
    style.map("Custom.Treeview.Heading",
              background=[('active', 'neon green'), ('hover', 'neon green')],
              foreground=[('active', 'white'), ('hover', 'white')])
    style.configure("Custom.Treeview", highlightthickness=1, bd=1, relief="solid")
    
    style.configure("TButton",
                    font=("Helvetica", 10),
                    background="light blue",
                    foreground="black")
    style.map("TButton",
              background=[('active', 'neon green'), ('hover', 'neon green')],
              foreground=[('active', 'black'), ('hover', 'black')])

def create_refresh_button(root, command):
    button = ttk.Button(root, text="Refresh", command=command, style="TButton")
    button.pack(side=tk.TOP, pady=10)
    return button

def create_percentage_label(root):
    percentage_label = ttk.Label(root, text="0%", font=("Helvetica", 10))
    percentage_label.pack(side=tk.TOP, padx=10, pady=5)
    return percentage_label
