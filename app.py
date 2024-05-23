import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import os

# Directories containing the Excel files
directories = {
    "GT 5": r"D:\Python Programming\GT Operating parameters\Operating Parameters\GT 5",
    "GT 6": r"D:\Python Programming\GT Operating parameters\Operating Parameters\GT 6",
    "GT 7": r"D:\Python Programming\GT Operating parameters\Operating Parameters\GT 7"
}

def load_data(directory):
    # Get a list of all Excel files in the directory
    file_list = [file for file in os.listdir(directory) if file.endswith(".XLSX")]

    # Read data from all Excel files and concatenate into a single DataFrame
    dfs = []
    for file in file_list:
        file_path = os.path.join(directory, file)
        df = pd.read_excel(file_path)

        # Convert "Measurement time" column to datetime type
        df["Measurement time"] = pd.to_datetime(df["Measurement time"], format="%H:%M:%S").dt.time

        # Convert "Date" column to string
        df["Date"] = df["Date"].astype(str)

        # Concatenate "Date" and "Measurement time" to create a new datetime column
        df["Datetime"] = pd.to_datetime(df["Date"] + " " + df["Measurement time"].astype(str))

        dfs.append(df)

    # Concatenate all DataFrames
    df = pd.concat(dfs, ignore_index=True)
    return df

# Function to fill missing hours and interpolate
def fill_missing_hours_and_interpolate(df):
    df = df.drop_duplicates(subset='Datetime')
    df = df.set_index('Datetime').resample('h').asfreq().infer_objects().reset_index()  # infer_objects here
    numeric_columns = df.select_dtypes(include='number').columns  # Select numeric columns for interpolation
    df[numeric_columns] = df[numeric_columns].interpolate('linear', dtype='float64')  # Perform interpolation
    return df

# Create the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Define the layout of the app
app.layout = dbc.Container(
    [
        dbc.Row(
            dbc.Col(
                html.H1("Gas Turbine Operating Parameters Dashboard", className="text-center my-4")
            )
        ),
        dbc.Row(
            [
                dbc.Col(
                    dbc.Card(
                        dbc.CardBody(
                            [
                                html.H5("Select Unit", className="card-title"),
                                dcc.RadioItems(
                                    id="directory-radio",
                                    options=[{"label": k, "value": k} for k in directories.keys()],
                                    value="GT 5",  # Set default value
                                    labelStyle={'display': 'inline-block', 'margin-right': '20px'},
                                    style={"margin-right": "10px"},  # Add margin between radio buttons and labels
                                    className="form-check-inline"
                                ),
                            ]
                        ),
                        className="h-100"  # Ensure the card takes full height
                    ),
                    width={"size": 3, "offset": 1},  # Shift the card a bit to the right
                ),
                dbc.Col(
                    dbc.Card(
                        dbc.CardBody(
                            [
                                html.H5("Select Measurement Point", className="card-title"),
                                dcc.Dropdown(id="description-dropdown", className="form-control"),
                            ]
                        ),
                        className="h-100"  # Ensure the card takes full height
                    ),
                    width=5,
                    className="offset-md-2"  # Adjust the right side offset to balance the layout
                ),
            ],
            className="mb-4"
        ),
        dbc.Row(
            dbc.Col(
                dbc.Card(
                    dbc.CardBody(
                        dcc.Loading(
                            id="loading-graph",
                            type="circle",
                            children=dcc.Graph(id="data-plot")
                        )
                    )
                )
            )
        ),
    ],
    fluid=True,
    style={"backgroundColor": "#f8f9fa"}  # Set the background color here
)

# Combined callback to update dropdown options and data plot based on inputs
@app.callback(
    [Output("description-dropdown", "options"),
     Output("description-dropdown", "value"),
     Output("data-plot", "figure")],
    [Input("directory-radio", "value"),
     Input("description-dropdown", "value")]
)
def update_output(directory_name, selected_description):
    ctx = dash.callback_context

    # Determine which input triggered the callback
    triggered_input = ctx.triggered[0]["prop_id"].split(".")[0]

    # Load data from the selected directory
    directory = directories[directory_name]
    df = load_data(directory)

    # Generate options for the dropdown
    description_options = [
        {"label": desc, "value": desc}
        for desc in df["Description of measuring point"].unique()
    ]

    # Default description value
    if triggered_input == "directory-radio" or not selected_description:
        selected_description = df["Description of measuring point"].unique()[0]

    # Filter data for the selected description
    filtered_df = df[df["Description of measuring point"] == selected_description]

    # Fill missing hours and interpolate
    connected_df = fill_missing_hours_and_interpolate(filtered_df)

    # Plot the data with enhanced styling
    fig = px.line(
        connected_df,
        x="Datetime",
        y="Meas/TotCountrRdg   _",
        title=f"Data for {selected_description} ({connected_df['CharactstcUnit'].iloc[0]})",
        template="plotly_white",
        line_shape='spline'  # Smooth lines
    )

    fig.update_xaxes(
        title="Measurement Time",
        tickangle=45,
        rangeslider_visible=True,
        showgrid=True,
        gridcolor='LightGrey',
        rangeselector=dict(
            buttons=list(
                [
                    dict(count=1, label="1d", step="day", stepmode="backward"),
                    dict(count=1, label="1m", step="month", stepmode="backward"),
                    dict(count=6, label="6m", step="month", stepmode="backward"),
                    dict(count=1, label="YTD", step="year", stepmode="todate"),
                    dict(count=1, label="1y", step="year", stepmode="backward"),
                    dict(step="all"),
                ]
            )
        )
    )

    fig.update_yaxes(
        title="Measurement Value",
        showgrid=True,
        gridcolor='LightGrey'
    )

    fig.update_layout(
        title={
            'text': f"Data for {selected_description} ({connected_df['CharactstcUnit'].iloc[0]})",
            'y':0.95,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=14,
            color="Black"
        ),
        plot_bgcolor='rgba(240,240,240,1)',
        paper_bgcolor='rgba(248,249,250,1)',
        margin=dict(l=20, r=20, t=50, b=20)
    )

    fig.update_traces(connectgaps=True)  # Connect gaps in the data

    return description_options, selected_description, fig

# Run the app
if __name__ == "__main__":
    app.run_server(debug=True)
