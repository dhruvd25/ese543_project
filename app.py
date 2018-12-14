import dash
import dash_core_components as dcc
import dash_html_components as html
import pandas as pd
import dash_table_experiments as dt
import zipfile

# Load in data set
zip = zipfile.ZipFile(r'output.xlsx.zip?raw=true')  
zip.extractall()  

sheet_to_df_map = pd.ExcelFile('output.xlsx')
dropdown_options = pd.read_excel('output.xlsx', sheet_name=None)

# Create the dash app
app = dash.Dash()

# Define the layout for the drop down menu
app.layout = html.Div([
    html.H2("Select Sheet Number"),
    html.Div([dcc.Dropdown(id="field_dropdown", options=[{
                               'label': i,
                               'value': i
                           } for i in dropdown_options],
                           value='Sheet3')],
             style={'width': '25%',
                    'display': 'inline-block'}),
    dt.DataTable(rows=[{}],
                 row_selectable=True,
                 filterable=True,
                 sortable=True,
                 selected_row_indices=[],
                 id='datatable')
])


@app.callback(
    dash.dependencies.Output('datatable', 'rows'),
    [dash.dependencies.Input('field_dropdown', 'value')])
def update_datatable(user_selection):
    if user_selection == 'Sheet1':
        return sheet_to_df_map.parse(0).to_dict('records')
    elif user_selection == 'Sheet2':
        return sheet_to_df_map.parse(1).to_dict('records')
    else:
        return sheet_to_df_map.parse(2).to_dict('records')


if __name__ == '__main__':
    app.run_server()