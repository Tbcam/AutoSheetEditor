import tkinter as tk
import os
from datetime import datetime
from tkinter import filedialog
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter

try:
    root = tk.Tk()
    root.withdraw()

    print('Choose the file type to combine:')
    file_type = input('Enter "excel" for Excel files or "pdf" for PDF files: ').strip().lower()

    if file_type == 'excel':
        # Select Excel files
        print('Select your Excel files')
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        print(file_paths)

        # List to hold DataFrames
        df_list = []

        for file_path in file_paths:
            print('\nReading file: ' + file_path)

            # Read the Excel file
            df = pd.read_excel(file_path)

            # Add a new column with the name of the Excel file
            filename = os.path.basename(file_path)
            df['Source File'] = filename

            # Append DataFrame to list
            df_list.append(df)

        # Concatenate all DataFrames in the list
        combined_df = pd.concat(df_list, ignore_index=True)

        # Make working folder for combined file
        current_datetime = datetime.now()
        date_str = current_datetime.strftime("%Y-%m-%d")
        time_str = current_datetime.strftime("%H-%M-%S")
        combined_folder_name = f"Combined_{date_str}_{time_str}"
        combined_folder_path = os.path.join(os.path.dirname(file_paths[0]), combined_folder_name)
        if not os.path.exists(combined_folder_path):
            os.makedirs(combined_folder_path)

        # Save the combined DataFrame as a new Excel file
        combined_file_path = os.path.join(combined_folder_path, "combined_file.xlsx")
        combined_df.to_excel(combined_file_path, index=False)

        print(f'Combined Excel file is saved to: {combined_file_path}')

    elif file_type == 'pdf':
        # Select PDF files
        print('Select your PDF files')
        file_paths = filedialog.askopenfilenames(filetypes=[('PDF Files', '*.pdf')])
        print(file_paths)

        # PDF writer to hold combined content
        pdf_writer = PdfWriter()

        for file_path in file_paths:
            print('\nReading file: ' + file_path)

            # Read the PDF file
            pdf_reader = PdfReader(file_path)

            # Add all pages to the writer
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

        # Make working folder for combined file
        current_datetime = datetime.now()
        date_str = current_datetime.strftime("%Y-%m-%d")
        time_str = current_datetime.strftime("%H-%M-%S")
        combined_folder_name = f"Combined_{date_str}_{time_str}"
        combined_folder_path = os.path.join(os.path.dirname(file_paths[0]), combined_folder_name)
        if not os.path.exists(combined_folder_path):
            os.makedirs(combined_folder_path)

        # Save the combined PDF as a new file
        combined_file_path = os.path.join(combined_folder_path, "combined_file.pdf")
        with open(combined_file_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

        print(f'Combined PDF file is saved to: {combined_file_path}')

    else:
        print('Invalid file type selected. Please run the script again and choose either "excel" or "pdf".')

    root.destroy()
except Exception as e:
    print(f'\nERROR: Combining files failed. An exception occurred: {str(e)}')

input('\nPress enter to exit')
