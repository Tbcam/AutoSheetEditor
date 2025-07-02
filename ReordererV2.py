import tkinter as tk
import os
from datetime import datetime
from tkinter import filedialog
import pandas as pd
from pathlib import Path

try:
    
    root = tk.Tk()
    root.withdraw()

    # base data file path
    print('Select your base file')
    file_path = filedialog.askopenfilename()
    print(file_path)


    # make working folder
    current_datetime = datetime.now()
    date_str = current_datetime.strftime("%Y-%m-%d")
    time_str = current_datetime.strftime("%H-%M-%S")
    folder_name = os.path.join(os.path.dirname(file_path), f"{date_str}_{time_str}")
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)


    # unedited data file path
    print('Select the file you want to edit')
    file_unedited_path = filedialog.askopenfilenames()
    print(file_unedited_path)

    count = 0

    for i in file_unedited_path:
        print('\nEditing this file: ' + file_unedited_path[count])

        # determine file format
        if i.endswith('.csv'):
            file_format = 'csv'
        elif i.endswith('.xlsx') or i.endswith('.xls'):
            file_format = 'excel'
        else:
            raise ValueError('This program only works with CSV and Excel files.')

        # read the base and unedited data files
        if file_format == 'csv':
            dfbase = pd.read_csv(file_path)
            dfunedited = pd.read_csv(i)
        elif file_format == 'excel':
            dfbase = pd.read_excel(file_path)
            dfunedited = pd.read_excel(i)

        # empty column deleter
        nan_value = float("NaN")
        dfbase.dropna(how='all', axis=1, inplace=True)

        # get column heads
        colbase = list(dfbase)
        missing_columns = [col for col in dfbase.columns if col not in dfunedited.columns]
        if len(missing_columns) == 0:
            print('\nSuccess no missing columns has been found.')
        else:
            print('\n This file. ' + file_unedited_path[count] +
                  'had these columns missing: ' + ' '.join(missing_columns))

        # apply heads
        dfunedited = dfunedited.reindex(columns=colbase)
        dfafter = dfunedited[colbase]
        if count == 0:
            print('\nThese base column heads has been used:\n', colbase)
        else:
            print('\nSame base columns heads has been used for editing.')

        print('\nFile is exported to: ' + folder_name)

        # Get file names
        dfafter_name = Path(i).name

        # Saves working folder
        save_path = os.path.join(folder_name, dfafter_name)
        if file_format == 'csv':
            dfafter.to_csv(save_path, index=False)
        elif file_format == 'excel':
            dfafter.to_excel(save_path, index=False)
        count = count + 1

    root.clipboard_clear()  # clear clipboard contents first
    root.clipboard_append(os.getcwd() + '/' + folder_name)
    root.destroy()
except Exception as e:
    print(f'\nERROR: Editing failed an exception occurred: {str(e)}')
    print('There might be an error in this file: ' + file_unedited_path[count])

input('\nPress enter to exit')
