# img_viewer.py

import PySimpleGUI as sg
import os.path
import pandas as pd

# software design setting
_version="v0.3"
textBock_size=75
gap_size=5
outputDisplay_size=45

#set Debug mode
DEBUG=True

# initial parameter setting
filter_column='LAST STATUS'
result_column='CON#'
fkey='SIP-L'

# Global variable
filter_list=[]
result_list=[]

def select_update(data_list):
    try:
        text = '\n'.join(item[:] for item in data_list) 
        window["-DATA-"].update(text)
    except:
        pass
def auto_update(twod_list):
    try:
        text = '\n'.join( [j for sub in twod_list for j in sub])
        print(text)

        window["-DATA-"].update(text)
    except:
        pass

select_folder_column = [
    [
        sg.Text("Folder:", size=(gap_size, 1)),
        sg.In(
               size=(textBock_size, 1), enable_events=True, key="-FOLDER-"
            ),
        sg.FolderBrowse(),
    ],
    [
        sg.Text("Filter:", size=(gap_size, 1)),
        sg.Combo(
          values=[], size=(textBock_size, 1), expand_x=False, enable_events=True,  
          readonly=False, key='-FILTER-', 
          default_value=fkey
        ),
    ],
    [
        sg.Text("Result:", size=(gap_size, 1)),
        sg.InputText(
                default_text=result_column, size=(textBock_size, 1), enable_events=True, key="-RESULT-"
            ),
        # sg.In(size=(textBock_size, 1), enable_events=True, key="-RESULT-"),
        # sg.Combo(
        #   values=[], size=(textBock_size, 1), expand_x=False, enable_events=True,  
        #   readonly=False, key='-RESULT-', 
        #   default_value=result_column
        # ),
    ],
    [
        sg.Text("Choose an Excel file from list at the bottom:")
    ],
]

# First the window layout in 2 columns
file_list_column = [
    [
        sg.Listbox(
            values=[], enable_events=True, size=(outputDisplay_size, 20), key="-FILE LIST-"
        ),
    ],
]

# For now will only show the name of the file that was chosen
data_viewer_column = [
    [sg.Multiline(
        default_text=[], size=(outputDisplay_size, 20), key='-DATA-',
        # disabled=True,
        do_not_clear=True,
        ),
    ],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(select_folder_column),
    ],
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(data_viewer_column),
    ]
]

window = sg.Window("Excel Data Filter GUI "+_version, layout)
while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    
    # Folder name was filled in, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []

        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            # and f.lower().endswith((".xls", ".xlsx"))
            and f.lower().endswith((".xlsx", ".xls"))
        ]
        window["-FILE LIST-"].update(fnames)

    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
            df = pd.read_excel(open(filename, 'rb'))

            if DEBUG:
                print(filename)
                print(df)

            FILTER=df[filter_column]
            for x in FILTER:
                if x not in filter_list:
                    filter_list.append(x)
            print("FILTER::  ",filter_list)
            window["-FILTER-"].update(values=filter_list)
            window["-RESULT-"].update(result_column)

            # initial data by fkey
            key_list=[]

            for key in filter_list:
                try:
                    key.index(fkey)
                except ValueError:
                    print("Not found!")
                else:
                    # print(key)

                    filter_sel = df[df[filter_column] == key]
                    data_list = filter_sel[result_column].tolist()
                    key_list.append(data_list)
                    # extrack_list(key_list)


            # filter_sel = df[df[filter_column] == savekey]
            print(key_list)

            auto_update(key_list)
        except:
            pass
    # elif event == "-FILTER-" or event == "-RESULT-":
    elif event == "-FILTER-":
        filter_sel = df[df[filter_column] == values["-FILTER-"]]
        data_list = filter_sel[result_column].tolist()
        select_update(data_list)


window.close()

