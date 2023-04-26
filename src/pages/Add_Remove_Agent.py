from dash import html, dcc
import dash_bootstrap_components as dbc
from datetime import datetime
import pandas as pd
# noinspection PyCompatibility
import pathlib

PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("../").resolve()

# Import data from excel files
dt = pd.read_excel(DATA_PATH.joinpath("Data.xlsx"))
df = pd.read_excel(DATA_PATH.joinpath("Agents.xlsx"))

# Assign list of names
df['Name'] = df['Name'].str.upper()
name_list = df['Name']
first_name_list = df.loc[0, 'Name']

# Establishes timestamp
now = datetime.now()
timestamp_num = datetime.timestamp(now)
timestamp = datetime.fromtimestamp(timestamp_num)

# Define the page layout
layout = dbc.Container([
    dbc.Row([
        html.Center(
            html.H1("Add/Remove Agent", style={'color': '#0cc4db'})
        ),
        html.Br(),
        dbc.Row(
            dbc.Col(
                html.Hr(
                    style={"borderWidth": "0.5vh", "width": "121%", "borderColor": "#8888fd", "opacity": "unset", }),
                width={"size": 10, "offset": 0},
            ),
        ),
        dbc.Col([
            html.P("Select the name of the Agent you want to Add", style={'color': '#0cc4db'}),

            # Add agent
            html.Div(
                dcc.Input('', id='new-name-input', type='text')
            ),
            dbc.Button('Submit', color="primary", id='add-agent-btn', n_clicks=0),
            html.Div(id='add-agent-output', children='Enter a Name and press submit', style={'color': '#ffffff'}),
        ]),
        dbc.Col([
            # Remove agent
            html.P("Select the name of the Agent you want to Remove", style={'color': '#0cc4db'}),
            html.Div(
                dcc.Dropdown(name_list, first_name_list, id='remove-agent-dropdown', clearable=False)
            ),
            html.Div(id='remove-agent-output', style={'color': '#ffffff'}),
            dcc.ConfirmDialogProvider(children=dbc.Button('Remove'), id='remove-agent-btn',
                                      message='Are you sure you want to delete this agent?'
                                      ),
            html.Hr(style={"margin-bottom": "700px"}),

        ])
    ])
])
