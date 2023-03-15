import PySimpleGUI as sg 
import pandas as pd

#add some color to window
sg.theme('DarkTeal9')

#Path to excel file
EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('Favorite Color', size=(15,1)), sg.Combo(['Red', 'Black', 'Pink'], key='Favorite Color')],
    [sg.Text('I speak', size=(15,1)),
                        sg.Checkbox('English', key='English'),
                        sg.Checkbox('Bangla', key='Bangla'),
                        sg.Checkbox('French', key='French')],
    [sg.Text('Age', size=(15,1)), sg.Spin([i for i in range(1,30)],
                                          initial_value=1, key='Age')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple Data Entry Form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        print(event, values)
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()

window.close()
