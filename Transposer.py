import tkinter as tk
import os
from datetime import datetime
from tkinter import filedialog
import pandas as pd
from pathlib import Path

try:
    root = tk.Tk()
    root.withdraw()

    # Select Excel files
    print('Select your Excel files')
    file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
    print(file_paths)

    # Make working folder for transposed files
    current_datetime = datetime.now()
    date_str = current_datetime.strftime("%Y-%m-%d")
    time_str = current_datetime.strftime("%H-%M-%S")
    transposed_folder_name = f"Transposed_{date_str}_{time_str}"
    transposed_folder_path = os.path.join(os.path.dirname(file_paths[0]), transposed_folder_name)
    if not os.path.exists(transposed_folder_path):
        os.makedirs(transposed_folder_path)

    for file_path in file_paths:
        print('\nTransposing this file: ' + file_path)

        # Read the Excel file
        df = pd.read_excel(file_path)

        # Transpose the DataFrame and reset the index
        df_transposed = df.transpose()

        # Set new header using the first row
        df_transposed.columns = df_transposed.iloc[0]
        df_transposed = df_transposed[1:]

        # Get the filename
        filename = Path(file_path).stem

        # Add a new column with the name of the Excel file
        df_transposed['Folder Name'] = filename

        # Save the transposed DataFrame as a new Excel file in the transposed folder
        save_path = os.path.join(transposed_folder_path, f"{filename}_transposed.xlsx")
        df_transposed.to_excel(save_path, index=False)

        print(f'Transposed file is saved to: {save_path}')

    root.destroy()
except Exception as e:
    print(f'\nERROR: Transposing failed. An exception occurred: {str(e)}')

input('\nPress enter to exit')
