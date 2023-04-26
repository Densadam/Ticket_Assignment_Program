from dash import html, dcc
import dash_bootstrap_components as dbc

image_path = "/assets/Logo.png"


# Define the navbar structure
def navbar():
    layout = html.Div([
        dbc.NavbarSimple(
            children=[
                dcc.Interval(id="time-interval", interval=60000),
                dbc.NavItem(dbc.NavLink("Tracker", href="/Tracker")),
                dbc.NavItem(dbc.NavLink("Add/Remove Agent", href="/Add_Remove_Agent")),
                dbc.NavItem(dbc.NavLink("Analytics", href="/Analytics")),
                html.Img(src=image_path),
                html.Div(id="current-time-output", style={'color': '#0cc4db'}),
            ],
            brand="Ticket Assignment Program - Created by Adam Densley",
            brand_href="/Tracker",
            color="#008dcf",
            dark=True,
        ),
    ])

    return layout
