# Upload File

import tkinter as tk
from tkinter import filedialog
import openpyxl


def upload_excel():
    # Specify your desired file name
    file_name = "PythonTestUpload.xlsx"

    # Get the directory path where the file will be saved
    save_directory = filedialog.askdirectory()

    # Construct the complete file path
    file_path = save_directory + "/" + file_name

    workbook = openpyxl.load_workbook(file_path)
    # Process the uploaded Excel file as needed
    workbook.close()
    print("File uploaded successfully.")


# Create a Tkinter window
window = tk.Tk()

# Create a button widget
upload_button = tk.Button(window, text="Upload Excel", command=upload_excel)
upload_button.pack()

# Run the Tkinter event loop
window.mainloop()
