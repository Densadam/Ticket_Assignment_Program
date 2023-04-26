from dash import html, dcc
import dash_bootstrap_components as dbc
import pandas as pd
# noinspection PyCompatibility
import pathlib

# Define path library for opening xslx data files
PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("../").resolve()

# Import data from excel files
dt = pd.read_excel(DATA_PATH.joinpath("Data.xlsx"))
df = pd.read_excel(DATA_PATH.joinpath("Agents.xlsx"))

df['Name'] = df['Name'].str.upper()
name_list = df['Name']

first_name_list = df.loc[0, 'Name']

# Define the Tracker page layout
layout = dbc.Container([

    dbc.Row([
        html.Center(
            html.H1("Ticket Tracker", style={'color': '#0cc4db'})
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

            # Working Agents List
            html.P(" - Agents Working Today - ",
                   style={'color': '#0cc4db', "font-weight": "bold", 'textAlign': 'center', 'font-size': '120%'}),

            dbc.Row([
                dcc.Dropdown(options=name_list, id='working-list', multi=True, persistence=True,
                             persistence_type='session'),
                dbc.Button("Set Working List", style={"width": "30%"}, color="primary", id='working-btn', n_clicks=0),
            ], justify="center"),

            html.Br(),
            html.Div(id='working-msg-output', style={'textAlign': 'center', 'color': '#ffffff'}),

            html.Hr(style={"borderWidth": "0.5vh", "width": "100%", "borderColor": "#8888fd", "opacity": "unset", }
                    ),

            # Auto-Assign Ticket Button
            html.Div('-Assign Ticket-',
                     style={'color': '#0cc4db', "font-weight": "bold", 'textAlign': 'center', 'font-size': '120%'}),
            html.Div(id='assignment-status-message', style={'textAlign': 'center', 'color': '#ffffff'}),
            html.Br(),
            dbc.Row(
                dbc.Button("Auto-Assign Ticket", style={"width": "30%"}, color="primary", id='auto-assign-btn',
                           n_clicks=0),
                justify="center"),

            # Manually assign ticket
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.2vh", "width": "120%", "borderColor": "#0cc4db",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0},
                ),
            ),
            dbc.Row([
                dbc.Button("Manually Assign Ticket", style={"width": "40%"}, color="success", id='manual-assign-btn',
                           n_clicks=0),
                dcc.Dropdown(name_list, first_name_list, id='assign_name_dropdown', clearable=False,
                             className="text-success", style={'textAlign': 'center', 'width': '60%'}),
            ], align="center", justify="center"),
        ]),

        dbc.Col([
            html.Br(),
            # Ticket information display
            html.P("- Ticket Information -",
                   style={'color': '#0cc4db', 'textAlign': 'center', "font-weight": "bold", 'font-size': '120%'}),
            html.Br(),
            html.Tr([
                html.Td('Current Ticket is assigned to: ', style={'color': '#8cff9a'}),
                html.Td(id='assignee-output', style={'color': '#8cff9a'})
            ]),
            html.Br(),
            html.Tr([
                html.Td('Next Ticket will be assigned to: ', style={'color': '#fbfeac'}),
                html.Td(id='next-assignee-output', style={'color': '#fbfeac'})
            ]),
            html.Br(),
            html.Tr([
                html.Td(children='Previous Ticket was assigned to: ', style={'color': '#76cffd'}),
                html.Td(id='prev-assignee-output', style={'color': '#76cffd'})
            ]),
            html.Br(),
            html.Tr([
                html.Td('Tickets assigned today = ', style={'color': '#ffffff'}),
                html.Td(id='tickets-today-count-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Total Tickets worked = ', style={'color': '#ffffff'}),
                html.Td(id='total-count-output', style={'color': '#ffffff'})
            ]),
            html.Br(),
            html.Br(),

            # Ticket Undo button
            dcc.ConfirmDialogProvider(children=dbc.Button('Undo Last Ticket Assignment', outline=True, color="danger",
                                                          style={"width": "40%"}),
                                      id='undo-btn', message='Are you sure you want to undo the last Ticket assigment?'
                                      ),
            html.Br(),
            html.Br(),
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#8888fd",
                                   "opacity": "unset", "margin-bottom": "450px"}), width={"size": 10, "offset": 0},
                ),
            ),
        ])
    ])
])
