import dash
from dash import html, ctx, dcc, State
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from time import sleep
from datetime import datetime
import pandas as pd
import pytz

# Connect to main app.py file
from app import app, server

# Connect to pages folder and the pages apps inside it
from pages import Tracker, Add_Remove_Agent, Analytics

# Connect the navbar to the index
from components import navbar

# define the navbar
nav = navbar.navbar()

# Define the index page layout
app.layout = dbc.Container(html.Div(
    children=[

        # Core Homepage
        dcc.Location(id='url', refresh=True),
        nav,
        html.Div(id='page-content', children=[], style={'backgroundColor': '#4c4c4c', 'opacity': '0.95'}),

        # dcc.Store inside the user's current browser session
        dcc.Store(id='store-assignment-info-data', data=[], storage_type='memory'),  # 'memory', 'local' or 'session'

    ]),
    style={
        'background-image': 'url("/assets/Background.png")',
        'background-repeat': 'no-repeat',
        'background-position': 'center center',
        'background-size': '850px 800px'
    },
    fluid=True,
    className="stylesheets"
)


# Sets up paths for each page
@app.callback(Output('page-content', 'children'),
              Input('url', 'pathname')
              )
def display_page(pathname):
    if pathname == '/Tracker':
        return Tracker.layout
    if pathname == '/Add_Remove_Agent':
        return Add_Remove_Agent.layout
    if pathname == '/Analytics':
        return Analytics.layout

    else:
        return html.Div("Please choose a link", style={'color': '#ffffff'}), html.Div(style={"margin-bottom": "875px"})


##########################
#### Add/Remove Agent ####
##########################


@app.callback(
    Output('add-agent-output', 'children'),
    Output('remove-agent-output', 'children'),
    Output('remove-agent-dropdown', 'options'),

    Input('add-agent-btn', 'n_clicks'),
    Input('remove-agent-btn', 'submit_n_clicks'),
    Input('remove-agent-dropdown', 'value'),
    State('remove-agent-dropdown', 'options'),
    State('new-name-input', 'value'),
    prevent_initial_call=False
)
#######################
#### Add New Agent ####
#######################

def update_add_remove_agent(add_agent_btn, remove_agent_btn, selected_name_value, dropdown_list_update, name_input):
    sleep(1)

    # Reads xsls file which contains agent data
    df = pd.read_excel(r'Agents.xlsx')

    # Sets baseline messages
    rmv_agent_msg = 'Select name of the agent who you want to remove from the list of agents'
    add_agent_msg = 'Select name of the agent who you want to add to the list of agents'

    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if triggered_id == "add-agent-btn":

        # checks if value is blank or has non-alpha characters
        if not name_input.isalpha():
            add_agent_msg = 'No spaces, numbers or special characters are allowed'

        else:
            value_name = name_input
            selected_name = value_name.upper()

            # Creates seperate list to compare if duplicates exist.
            name_list = df['Name']
            name_list.loc[len(name_list)] = selected_name

            # If duplicates exist on seperate list, error message generated. Else, add name to main list.
            if len(name_list) != len(set(name_list)):
                add_agent_msg = 'Duplicate name, please enter a different name!'
                dup_check = True
            else:
                dup_check = False

            # Checks if duplicates status is false, then adds name to Agents xslx file and creates individual agent file
            if not dup_check:

                add_agent_msg = 'Added {} to the list of agents!'.format(selected_name)

                add_name = [selected_name, 0, 1, False]
                df.loc[len(df)] = add_name

                # Gets time in EST and creates timestamp
                now = datetime.now()
                timestamp_num = datetime.timestamp(now)
                time_now = datetime.fromtimestamp(timestamp_num)
                timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

                # Creates individual xslx file for the agent
                writer = pd.ExcelWriter(selected_name + '.xlsx', engine='xlsxwriter')
                with pd.ExcelWriter(selected_name + '.xlsx') as writer:

                    agent_created = {selected_name + '_Worked': [timestamp], "Action": ['Agent Created']}
                    df_new_agent = pd.DataFrame(agent_created)
                    df_new_agent.to_excel(writer, sheet_name=selected_name + '_Worked', index=False)

                df['Name'] = df['Name'].str.upper()
                name_list = df['Name']

                df.to_excel('Agents.xlsx', index=False)

                dropdown_list_update = df['Name']

            else:
                pass

    ######################
    #### Remove Agent ####
    ######################

    if triggered_id == 'remove-agent-btn':

        if not selected_name_value:
            rmv_agent_msg = 'No agent selected'

        else:
            agent_count = df.shape[0]

            if agent_count == 1:
                rmv_agent_msg = 'You need at least one agent on the list'

            else:
                rmv_agent_msg = 'You have removed {} from the list of agents'.format(selected_name_value)
                df = df[df["Name"].str.contains(selected_name_value) == False]
                df.to_excel('Agents.xlsx', index=False)

    dropdown_list_update = df['Name']

    return add_agent_msg, rmv_agent_msg, dropdown_list_update


#################################
##### Change Working State ######
#################################

@app.callback(
    Output('store-assignment-info-data', 'data'),
    Output('working-msg-output', 'children'),
    Output('working-list', 'options'),

    Input('working-btn', 'n_clicks'),
    Input('working-list', 'value'),
    State('working-list', 'options'),
    prevent_initial_call=False
)
def update_working(set_working_button, working_list_value, dropdown_list_update):
    # Reads xsls files which contains long term stored data
    df = pd.read_excel(r'Agents.xlsx')
    dt = pd.read_excel(r'Data.xlsx')

    sleep(1)

    now = datetime.now()
    timestamp_num = datetime.timestamp(now)
    time_now = datetime.fromtimestamp(timestamp_num)

    # Sets up today's date and yesterday's data
    day = time_now.strftime("%d")
    today = day
    yesterday = dt.loc[0, 'Date']
    dt.loc[0, 'Date'] = today

    next_assignee = ' push Set Working List button to get next agent'

    working_names = pd.DataFrame(df.loc[df['Working'] == True, 'Name'])
    working_name_list = working_names['Name'].str.split()
    working_list_msg = 'List of agents working: ' + working_name_list.to_string(index=False)

    # checks if time is different day, if true, resets tickets worked to zero
    if int(today) != yesterday:

        df['Tickets_Worked'] = 0

    else:
        pass

    triggered_id = ctx.triggered_id

    if triggered_id == 'working-btn':

        # Checks if there are any workers selected
        if not working_list_value:

            df['Working'] = False
            working_list_msg = "There aren't any agents selected, please select an Agent first!"

        else:

            # Updates selected agent values and updates working list. 
            list_value = working_list_value
            df_working_list_value = pd.Series(list_value)
            working_value = df['Name'].isin(df_working_list_value)
            df['Working'] = working_value

            # Finds agent with the highest selection value and is working
            work_list = df[df['Working'] == True]

            df_agent_select = pd.DataFrame(work_list)
            max_work_index = df_agent_select['Agent_Selected'].idxmax()
            max_agent_name = df.at[max_work_index, 'Name']
            dt.at[0, 'Next_Agent'] = max_agent_name

            next_assignee = max_agent_name

            # Creates list of agents who are working
            df_working_list = pd.DataFrame(df.loc[df['Working'] == True, 'Name'])
            worklist_name = df_working_list['Name'].str.split()
            working_list_msg = 'List of agents working: ' + worklist_name.to_string(index=False)

    else:
        pass

    dropdown_list_update = df['Name']

    df.to_excel('Agents.xlsx', index=False)
    dt.to_excel('Data.xlsx', index=False)

    return next_assignee, working_list_msg, dropdown_list_update


###############################################
### Manual/Automatic/Undo Assign Tickets  ####
#############################################

@app.callback(
    # Auto/Mnl Ticket assign
    Output('assignment-status-message', 'children'),
    Output('total-count-output', 'children'),
    Output('assignee-output', 'children'),
    Output('tickets-today-count-output', 'children'),
    Output('prev-assignee-output', 'children'),
    Output('next-assignee-output', 'children'),
    Output('assign_name_dropdown', 'options'),

    Input('undo-btn', 'submit_n_clicks'),
    Input('auto-assign-btn', 'n_clicks'),
    Input('manual-assign-btn', 'n_clicks'),
    State('assign_name_dropdown', 'options'),
    State("assign_name_dropdown", "value"),
    Input('store-assignment-info-data', 'data'),

    prevent_initial_call=False
)
###########################################
#### Manual Ticket Assignment Button  ####
#########################################

# Manual assign ticket and setups button callback for Automatic and Undo ticket assignment
def update_ticket(submit_n_clicks, button2, button3, options, value, data):
    dt = pd.read_excel(r'Data.xlsx')
    df = pd.read_excel(r'Agents.xlsx')

    status_msg = "Either manually or automatically assign agent a ticket"

    min_work_index = df['Agent_Selected'].idxmin()
    min_agent_name = df.at[min_work_index, 'Name']
    assignee = '{} at {}'.format(min_agent_name, dt.at[0, 'Current_Time'])

    # Gets sum of total tickets worked for the day
    day_output = df['Tickets_Worked'].sum()

    total_output = dt['Total_Tickets']

    prev_assignee = dt.at[0, 'Prev_Agent']
    next_assignee = data

    # Sets button trigger ID
    triggered_id = ctx.triggered_id

    if triggered_id == 'manual-assign-btn':

        # Reads/creates dataframes from Excel files
        df = pd.read_excel(r'Agents.xlsx')
        dt = pd.read_excel(r'Data.xlsx')

        # Gets agent name from dropdown and finds index value
        selected_name = value
        selected_index = df[df['Name'] == selected_name].index.values

        # Gets sum of total tickets worked for the day
        day_output = df['Tickets_Worked'].sum()

        # Checks if selected agent is working today then returns True/False value
        working_true = df.at[int(selected_index), 'Working']

        # Checks if agent is working today value is False
        if not working_true:

            status_msg = "Are you sure you selected the right agent? This agent isn't set to work today!"
            assignee = "Agent not assigned"

            # Finds agent who has the lowest selection value (last agent who worked ticket or a new agent)
            prev_assignee = dt.at[0, 'Prev_Agent']
            next_assignee = dt.at[0, 'Next_Agent']

        else:

            # Gets the time
            now = datetime.now()
            timestamp_num = datetime.timestamp(now)
            time_now = datetime.fromtimestamp(timestamp_num)
            timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')
            time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

            dt['Current_Time'] = time_ez
            time_atm = dt.at[0, 'Current_Time']

            # Finds agent who has the lowest selection value (last agent who worked ticket or a new agent)
            min_work_index = df['Agent_Selected'].idxmin()
            min_agent_name = df.at[min_work_index, 'Name']
            prev_assignee = min_agent_name
            dt.at[0, 'Prev_Agent'] = prev_assignee

            # Adds +1 ticket worked to selected agent
            df.at[int(selected_index), 'Tickets_Worked'] += 1

            # Add +1 to selection count value to all agents
            df['Agent_Selected'] += 1

            # Resets selection value count of selected agent to zero
            df.at[int(selected_index), 'Agent_Selected'] = 0

            # Add +1 to total worked tickets
            dt['Total_Tickets'] += 1

            # Reads personal agent Excel files
            df_assign = pd.read_excel(selected_name + '.xlsx', sheet_name=selected_name + '_Worked')

            # Makes copy of agent data and adds new data
            df_assign_copy_a = pd.DataFrame(df_assign[[selected_name + '_Worked', 'Action']])
            df_assign_copy_b = pd.DataFrame(
                {selected_name + '_Worked': [timestamp], "Action": ['Manually Assigned Ticket']})
            df_assign = pd.concat([df_assign_copy_a, df_assign_copy_b])

            # writes data to excel file
            writer = pd.ExcelWriter(selected_name + '.xlsx', engine='xlsxwriter')
            with pd.ExcelWriter(selected_name + '.xlsx') as writer:
                df_assign.to_excel(writer, sheet_name=selected_name + '_Worked', index=False)

            # Assigns value of total tickets worked
            total_output = dt['Total_Tickets']

            # Gets sum of total tickets worked for the day
            day_output = df['Tickets_Worked'].sum()

            # Returns message of changes made
            status_msg = '{} was manually assigned the next Ticket at {}'.format(selected_name, time_ez)
            assignee = '{} at {}'.format(selected_name, time_atm)

            # Finds agent with the highest selection value and is working
            work_list = df[df['Working'] == True]
            df_working_list = pd.DataFrame(work_list)
            max_work_index = df_working_list['Agent_Selected'].idxmax()
            max_agent_name = df.at[max_work_index, 'Name']
            next_assignee = max_agent_name
            dt.at[0, 'Next_Agent'] = next_assignee

            # Writes data to Excel files
            df.to_excel('Agents.xlsx', index=False)
            dt.to_excel('Data.xlsx', index=False)

    # calls on auto_assign assign button function if pressed
    if triggered_id == 'auto-assign-btn':
        return auto_assign()

    elif triggered_id == 'undo-btn':
        return undo_assign()
    else:
        pass
    options = df['Name']

    return status_msg, total_output, assignee, day_output, prev_assignee, next_assignee, options


#############################################
#### Automatic Ticket Assignment Button  ####
#############################################

# Auto Assign ticket based on a selection value counter. The agent with the highest value is selected for next ticket.
# Then selection value is reset for selected agent and all other agents gain +1 to selection value.
def auto_assign():
    df = pd.read_excel(r'Agents.xlsx')
    dt = pd.read_excel(r'Data.xlsx')

    # Checks if any agents are working
    if df['Working'].values.sum() != 0:

        # Gets the time
        now = datetime.now()
        timestamp_num = datetime.timestamp(now)
        time_now = datetime.fromtimestamp(timestamp_num)
        timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')
        time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

        dt['Current_Time'] = time_ez
        time_atm = dt.at[0, 'Current_Time']

        # Creates list of agents who are working today
        work_list = df[df['Working'] == True]

        # Finds agent who has the lowest selection value (last agent who worked ticket or a new agent)
        min_work_index = df['Agent_Selected'].idxmin()
        min_agent_name = df.at[min_work_index, 'Name']
        prev_assignee = min_agent_name
        dt.at[0, 'Prev_Agent'] = prev_assignee

        # Creates Data Frame of agents who are working today, then finds the agent with highest slection count value
        df_working_list = pd.DataFrame(work_list)
        max_work_index = df_working_list['Agent_Selected'].idxmax()
        max_agent_name = df.at[max_work_index, 'Name']

        # Reads and creates Data Frame of personal agent Excel file
        df_assign = pd.read_excel(max_agent_name + '.xlsx', sheet_name=max_agent_name + '_Worked')

        # Adds +1 count to all agent's selection count value
        df['Agent_Selected'] += 1

        # Adds +1 total number of worked tickets
        df.at[max_work_index, 'Tickets_Worked'] += 1

        # Gets name of agent with the highest selection count value then returns it in message
        assignee_name = df.at[max_work_index, 'Name']
        assignee = '{} at {}'.format(assignee_name, time_ez)
        status_msg = "{} was automatically assigned the next Ticket at {}".format(assignee_name, time_atm)

        # Resets value of the highest selection count agent to 0
        df.at[max_work_index, 'Agent_Selected'] = 0

        # Adds +1 to total ever worked tickets
        dt['Total_Tickets'] += 1

        # Gets sum of total tickets worked for the day
        day_output = df['Tickets_Worked'].sum()

        # Assigns value of total tickets worked
        total_output = dt['Total_Tickets']

        # Makes copy of selected agent data frame, writes timestamp to copy of data frame and then merges them together
        df_assign_copy_a = pd.DataFrame(df_assign[[max_agent_name + '_Worked', 'Action']])
        df_assign_copy_b = pd.DataFrame(
            {max_agent_name + '_Worked': [timestamp], "Action": ['Automatically Assigned Ticket']})
        df_assign = pd.concat([df_assign_copy_a, df_assign_copy_b])

        # writes data to agent's personal excel file
        writer = pd.ExcelWriter(max_agent_name + '.xlsx', engine='xlsxwriter')
        with pd.ExcelWriter(max_agent_name + '.xlsx') as writer:
            df_assign.to_excel(writer, sheet_name=max_agent_name + '_Worked', index=False)

        # Finds agent with the highest selection value and is working
        work_list = df[df['Working'] == True]
        df_working_list = pd.DataFrame(work_list)
        max_work_index = df_working_list['Agent_Selected'].idxmax()
        max_agent_name = df.at[max_work_index, 'Name']
        next_assignee = max_agent_name
        dt['Next_Agent'] = next_assignee

        # Saves changes to Excel files
        df.to_excel('Agents.xlsx', index=False)
        dt.to_excel('Data.xlsx', index=False)

    else:

        # Returns error messages
        status_msg = 'No agents selected to work today!'
        assignee = 'Agent not assigned'
        day_output = df['Tickets_Worked'].sum()

        total_output = dt['Total_Tickets']

        prev_assignee = dt.at[0, 'Prev_Agent']
        next_assignee = dt.at[0, 'Next_Agent']

    options = df['Name']

    return status_msg, total_output, assignee, day_output, prev_assignee, next_assignee, options


########################################
#### Undo Ticket Assignment Button  ####
########################################


# Undo last ticket assignment
def undo_assign():
    df = pd.read_excel(r'Agents.xlsx')
    dt = pd.read_excel(r'Data.xlsx')

    min_work_index = df['Agent_Selected'].idxmin()

    # Checks if any agents are working, if true: search agent with the lowest count then -= 1 to worked tickets.
    # Adds 9001 to count for assignee agent, then -= 1 count to all other agents.
    if df.at[min_work_index, 'Tickets_Worked'] == 0:

        time_atm = dt.at[0, 'Current_Time']

        status_msg = 'Unable to undo further'

        min_work_index = df['Agent_Selected'].idxmin()
        min_agent_name = df.at[min_work_index, 'Name']

        assignee = 'Unable to undo further'

        total_output = dt['Total_Tickets']

        day_output = df['Tickets_Worked'].sum()

        prev_assignee = dt.at[0, 'Prev_Agent']
        next_assignee = dt.at[0, 'Next_Agent']
    else:

        if df.at[min_work_index, 'Working']:

            # Gets the time
            now = datetime.now()
            timestamp_num = datetime.timestamp(now)
            time_now = datetime.fromtimestamp(timestamp_num)
            timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

            time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

            dt['Current_Time'] = time_ez
            time_atm = dt.at[0, 'Current_Time']

            status_msg = 'Click this button to undo last Ticket assignment'

            min_work_index = df['Agent_Selected'].idxmin()
            min_agent_name = df.at[min_work_index, 'Name']

            day_output = df['Tickets_Worked'].sum()

            df_assign = pd.read_excel(min_agent_name + '.xlsx', sheet_name=min_agent_name + '_Worked')

            df['Agent_Selected'] -= 1
            df.at[min_work_index, 'Tickets_Worked'] -= 1
            df.at[min_work_index, 'Agent_Selected'] = 9001
            dt['Total_Tickets'] -= 1

            total_output = dt['Total_Tickets']

            # Gets last agent who was assigned ticket and creates a dataframe
            df_last_agent = pd.DataFrame(df_assign[[min_agent_name + '_Worked', 'Action']])

            # Gets the index value of the last row which contains 'Assigned Ticket' value
            last_assign_ticket_index = \
                df_last_agent.loc[df_last_agent['Action'].str.contains('Assigned Ticket')].index[-1]

            # Replaces value of last "Assigned Ticket" with "Undone" statement and timestamp
            df_last_agent.iloc[last_assign_ticket_index, df_last_agent.columns.get_loc(
                'Action')] = 'ticket Assignment Undone @ ' + timestamp
            df_assign = df_last_agent

            # writes data to excel file
            writer = pd.ExcelWriter(min_agent_name + '.xlsx', engine='xlsxwriter')
            with pd.ExcelWriter(min_agent_name + '.xlsx') as writer:
                df_assign.to_excel(writer, sheet_name=min_agent_name + '_Worked', index=False)

            # Collects information numbers
            day_output = df['Tickets_Worked'].sum()

            # Finds agent with the highest selection value and is working
            work_list = df[df['Working'] == True]
            df_working_list = pd.DataFrame(work_list)
            max_work_index = df_working_list['Agent_Selected'].idxmax()
            max_agent_name = df.at[max_work_index, 'Name']
            next_assignee = max_agent_name

            min_work_index = df['Agent_Selected'].idxmin()
            min_agent_name_update = df.at[min_work_index, 'Name']
            dt['Prev_Agent'] = min_agent_name_update
            prev_assignee = min_agent_name_update

            assignee = '-Ticket Assignment Undone at- {}'.format(time_atm)

            status_msg = "{} was unassigned a Ticket at {}.".format(min_agent_name, time_atm)

            # Saves changes to Excel files
            df.to_excel('Agents.xlsx', index=False)
            dt.to_excel('Data.xlsx', index=False)

        else:

            min_work_index = df['Agent_Selected'].idxmin()
            min_agent_name_update = df.at[min_work_index, 'Name']
            dt['Prev_Agent'] = min_agent_name_update
            prev_assignee = min_agent_name_update

            status_msg = 'Please ensure {} is set to working today and try again!'.format(prev_assignee)

            total_output = dt['Total_Tickets']

            assignee = 'Error - Unable to Undo'

            day_output = df['Tickets_Worked'].sum()

            prev_assignee = dt.at[0, 'Prev_Agent']
            next_assignee = dt.at[0, 'Next_Agent']

    options = df['Name']

    return status_msg, total_output, assignee, day_output, prev_assignee, next_assignee, options


############################################################################
### Runs Report - Reads Excel Files and Combines into Master Dataframe #####
############################################################################

@app.callback(
    Output('report-output', 'children'),
    Output('bar-chart', 'figure'),
    Output('download_dropdown', 'options'),

    Input("run-report_btn", "n_clicks"),
    Input('date-picker-range', 'start_date'),
    Input('date-picker-range', 'end_date'),
    State('download_dropdown', 'options'),
    prevent_initial_call=True,
)
def run_report(n_clicks, start_date, end_date, dropdown_list_update):
    sleep(1)

    # Sets button trigger ID
    triggered_id = ctx.triggered_id

    if triggered_id == 'run-report_btn':

        # Reads Agent file in order to get list of agent names
        df = pd.read_excel(r'Agents.xlsx')
        agent_list = df['Name']

        # Sets up blank lists to be used later
        df_data_cmbnd = []

        # Cycles through list of agents and opens each Excel files, then merges into one DataFrame
        for i in agent_list:
            df_data = pd.read_excel(i + '.xlsx', sheet_name=[i + '_Worked'])
            df_data_worked = df_data.get(i + '_Worked')
            df_data_worked.rename(columns={i + '_Worked': "Date_Time"}, inplace=True)
            df_data_worked['Name'] = i
            df_data_cmbnd.append(df_data_worked)

        # Combined raw data files from eash agent into one master DataFrame
        df_date_time = pd.concat(df_data_cmbnd, ignore_index=True)

        # Converts DataFrame into Datetime object and sorts by date
        df_date_time['Date_Time'] = pd.to_datetime(df_date_time['Date_Time'])
        df_date_time.sort_values(by='Date_Time', ascending=True, inplace=True)
        df_date_time.reset_index(drop=True, inplace=True)
        df_master = pd.DataFrame(df_date_time)

        # Writes master DataFrame to excel file
        writer = pd.ExcelWriter('MASTER.xlsx', engine='xlsxwriter')
        with pd.ExcelWriter('MASTER.xlsx') as writer:
            df_master.to_excel(writer, sheet_name='MASTER_Worked', index=False)

        # Counts undo instances and returns value
        df_undo_count = df_master['Action'].str.count('Undone').sum()
        report_details = 'Finsihed report, there are {} counts of Tickets being undone'.format(df_undo_count)

        # Removes all rows that are not designated as assigning a ticket and resets index
        df_master = df_master[df_master["Action"].str.contains('Assigned Ticket')]
        df_master.reset_index(drop=True, inplace=True)

        # Convert timestamps to datetime object and sets the index
        df_master['Date_Time'] = pd.to_datetime(df_master['Date_Time'])
        df_master.set_index('Date_Time', inplace=True)

        # Sets up timestamp increments by day and counts instances
        daily_counts = df_master.resample('D').count()

        # Creates a mask filter for start/end date and creates bar graph based on filter
        mask = (daily_counts.index >= start_date) & (daily_counts.index <= end_date)
        filtered_counts = daily_counts.loc[mask]

        figure = {
            'data': [{

                'x': filtered_counts.index,
                'y': filtered_counts['Action'],

                'type': 'bar'
            }],
            'layout': {
                'title': 'Ticket Assignment History',
                'xaxis': {'title': 'Date'},
                'yaxis': {'title': 'Ticket Count'}
            }
        }

        # Adds names to dropdown list including Master option (Used for download option in next callback)
        add_mstr_value = 'MASTER'
        df.loc[df.index.max() + 1, 'Name'] = add_mstr_value
        dropdown_list_update = df['Name']

        return report_details, figure, dropdown_list_update


# Download callback
@app.callback(
    Output("download-dataframe-xlsx", "data"),
    Input("download_btn", "submit_n_clicks"),
    Input("download_dropdown", "value"),
    prevent_initial_call=True,
)
def download_xlsx(n_clicks, dnld_slcted_name):
    # Sets button trigger ID
    triggered_id = ctx.triggered_id

    # When button is pushed, it reads selected value then converts DataFrame into xslx and pushes download to user
    if triggered_id == 'download_btn':
        df_selected_dwnld = pd.read_excel(dnld_slcted_name + '.xlsx', sheet_name=dnld_slcted_name + '_Worked')
        df_dwnld = pd.DataFrame(df_selected_dwnld)
        return dcc.send_data_frame(df_dwnld.to_excel, dnld_slcted_name + '_Worked.xlsx', sheet_name="Tickets_Worked")

    # Used to give a cooldown time between button clicks and prevent abuse
    sleep(1)


@app.callback(
    Output("current-time-output", "children"),
    [Input("time-interval", "n_intervals")
     ])
def time_output(n):
    now = datetime.now()
    timestamp_num = datetime.timestamp(now)
    time_now = datetime.fromtimestamp(timestamp_num)
    current_time = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%m-%d-%Y %H:%M')
    time_message = current_time + 'EST'
    return time_message


# Run the app on localhost:8050
if __name__ == '__main__':
    app.run_server(debug=True)
