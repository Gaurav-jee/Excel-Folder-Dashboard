# 'C:/ART_GALLERYHPWHITE/2023/NEW BILLS'

import os
import pandas as pd
from datetime import datetime
import dash
import dash_bootstrap_components as dbc
from dash import dcc, html
from dash.dependencies import Input, Output
from flask import send_from_directory

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
                'Clickable Link': filename
            })
    return file_list

# Folder path (change this to your target folder)
folder_path = 'C:/ART_GALLERYHPWHITE/2023/NEW BILLS'

# Create a DataFrame
file_data = pd.DataFrame(get_file_details(folder_path))

# Initialize Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Flask server
server = app.server

# Route to serve files
@server.route("/files/<path:filename>")
def serve_file(filename):
    return send_from_directory(folder_path, filename)

# Layout of the dashboard
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("File Dashboard"), className="mb-4")
    ]),
    dbc.Row([
        dbc.Col(html.Table([
            html.Thead(
                html.Tr([html.Th(col) for col in file_data.columns])
            ),
            html.Tbody([
                html.Tr([
                    html.Td(file_data.iloc[i][col]) if col != 'Clickable Link' else html.Td(html.A('Open', href=f"/files/{file_data.iloc[i][col]}", target="_blank"))
                    for col in file_data.columns
                ]) for i in range(len(file_data))
            ])
        ], className="table table-striped table-bordered table-hover"), width=12)
    ])
], fluid=True)

# Run the app
if __name__ == "__main__":
    app.run_server(debug=True)
