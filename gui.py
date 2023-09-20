import PySimpleGUI as sg
import pandas as pd
import requests

# Set the GUI theme
sg.theme('DarkTeal9')

# Define the path to the Excel file
EXCEL_FILE = 'example.xlsx'
# Read existing data from the Excel file into a DataFrame
df = pd.read_excel(EXCEL_FILE)
# Define the Google Apps Script URL
GOOGLE_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx4v_doSPfcUEL2LMcM4ne5NYEvJ4lgjc55k9HPa6Y2Y74pRixRal38O7JtReCzxh1_5g/exec'

# Define the layout of the GUI
layout = [
    [sg.Text('Please fill up the following fields:  ')],
    [sg.Text('temi S/N', size=(15, 1)), sg.InputText(key='temi S/N')],
    [sg.Text('Item', size=(15, 1)), sg.InputText(key='Item')],
    [sg.Text('Old part', size=(15, 1)), sg.InputText(key='Old part')],
    [sg.Text('New part', size=(15, 1)), sg.InputText(key='New part')],
    [sg.Text('Issues', size=(15, 1)), sg.InputText(key='Issues')],
    [sg.Text('Date', size=(15, 1)), sg.InputText(key='Date')],
    [sg.Text('Remarks', size=(15, 1)), sg.Combo(['Warranty', 'Purchase', 'Office' + 'Spare'], key='Remarks')],
    [sg.Text('Customer', size=(15, 1)), sg.InputText(key='Customer')],
    [sg.Text('Quantity', size=(15, 1)), sg.Spin([i for i in range(0, 5)], initial_value=0, key='Quantity')],
    [sg.Submit(), sg.Exit()],
]

# Create the GUI window
window = sg.Window('simple data entry form', layout)



# Event loop to capture user interactions
while True:
    event, values = window.read()
    # Check if the window is closed or Exit button is clicked
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    # Check if the Submit button is clicked
    if event == 'Submit':
        # Create a dictionary with user-entered data
        new_data = {
            'temi S/N': values['temi S/N'],
            'Item': values['Item'],
            'Old part': values['Old part'],
            'New part': values['New part'],
            'Issues': values['Issues'],
            'Date': values['Date'],
            'Remarks': values['Remarks'],
            'Customer': values['Customer'],
            'Quantity': values['Quantity'],
        }
        
        # Create a new DataFrame from the dictionary
        new_df = pd.DataFrame(new_data, index=[len(df)])  # Use len(df) to set the index for a new row
        
        # Concatenate the new DataFrame with the existing DataFrame
        df = pd.concat([df, new_df], ignore_index=True)
        
        # Write the updated DataFrame to the Excel file
        df.to_excel(EXCEL_FILE, index=False)
    
        # Send the data to the Google Apps Script via HTTP POST
        response = requests.post(GOOGLE_APPS_SCRIPT_URL, data=new_data)

        if response.status_code == 200:
            sg.popup('Data Saved!')
        else:
            sg.popup('Data Saving Failed!')

# Close the GUI window when the loop ends
window.close()
