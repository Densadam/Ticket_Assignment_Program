from dash import html, dcc
import dash_bootstrap_components as dbc
from datetime import datetime
import pandas as pd
import pytz
# noinspection PyCompatibility
import pathlib

##########################
#### Intial report run ###
##########################
PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("../").resolve()

# Import data from excel files
df = pd.read_excel(DATA_PATH.joinpath("Agents.xlsx"))

add_mstr_value = 'MASTER'
df.loc[df.index.max() + 1, 'Name'] = add_mstr_value
name_list = df['Name']
first_name_list = df.loc[0, 'Name']

df_master = pd.read_excel(DATA_PATH.joinpath("MASTER.xlsx"), sheet_name='MASTER_Worked')

# Removes all rows that are not designated as assigning a ticket and resets index
df_master = df_master[df_master["Action"].str.contains('Assigned Ticket')]
df_master.reset_index(drop=True, inplace=True)

# Convert Ticket timestamps to datetime objects and sets the index as timestamps
df_master['Date_Time'] = pd.to_datetime(df_master['Date_Time'])
df_master.set_index('Date_Time', inplace=True)

# Sets up timestamp increments by day
daily_counts = df_master.resample('D').count()

# Gets current date in YYYY-MM-DD
now = datetime.now()
timestamp_num = datetime.timestamp(now)
time_now = datetime.fromtimestamp(timestamp_num)
time_now_date = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%Y-%m-%d')

layout = dbc.Container([

    html.Div(children=[

        # Run report button
        html.Br(),
        html.Center(
            dbc.Button("Run Report", style={"width": "30%", 'textAlign': 'center'}, color="primary",
                       id='run-report_btn',
                       n_clicks=0),
        ),

        html.Div(id='report-output', children='Click ^Run Report button^ to get started',
                 style={'color': '#0cc4db', 'textAlign': 'center', 'font-size': '120%'}
                 ),
        html.Br(),
        dbc.Col(
            html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#8888fd", "opacity": "unset", }
                    ),
            width={"size": 10, "offset": 0},
        ),

        # Ticket report Graph/Chart
        dcc.DatePickerRange(id='date-picker-range',
                            start_date=daily_counts.index.min().date(),
                            end_date=time_now_date),

        dcc.Graph(id='bar-chart'),
        html.Br(),
        html.Br(),
        html.Div(children="Select name of Agent who you want to download (MASTER is all Agent's data combined)",
                 style={'color': '#0cc4db', 'textAlign': 'center', 'font-size': '120%'}
                 ),

        # Agent data download
        html.Center(
            dcc.Dropdown(name_list, first_name_list, id='download_dropdown', clearable=False,
                         className="text-success", style={'textAlign': 'center',
                                                          'width': '55%',
                                                          'left': '10%',
                                                          'transform': 'translateX(8%)',
                                                          }
                         ),
        ),
        html.Br(),
        dcc.ConfirmDialogProvider(children=dbc.Button('Download Excel Data'), id="download_btn",
                                  message='Do you want to Download the selected Excel(.xsls) file?'
                                  ),
        dcc.Download(id="download-dataframe-xlsx"),
        html.Br(),
        dbc.Col(
            html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#8888fd", "opacity": "unset", }
                    ),
            width={"size": 10, "offset": 0},
        ),
        html.Br(),
        html.Br(),
        html.Br(),
    ], style={'textAlign': 'center'}
    ),
])
