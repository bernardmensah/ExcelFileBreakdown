import tkinter as tk
import pandas as pd
import os
from tkinter import filedialog


# Function to browse and select the Excel file to upload
def browse_file():
    file_path = filedialog.askopenfilename()
    # Update the file path label with the file path
    file_path_label.config(text=file_path)
    # Read the Excel file as a DataFrame
    df = pd.read_excel(file_path)

    return file_path, df


# Breakdown button
def breakdown_button_clicked():
    file_path, df = browse_file()
    breakdown(file_path, df)


# Function to break down the Excel file into smaller files
def breakdown(file_path, df):
    # Get the number of rows specified by the user
    num_rows = int(num_rows_entry.get())
    # Calculate the number of smaller files to create
    num_files = len(df) // num_rows
    if len(df) % num_rows != 0:
        num_files += 1
    # Create a list of the row indices to split the DataFrame at
    split_indices = [i*num_rows for i in range(num_files)]
    # Split the DataFrame into smaller DataFrames
    df_list = [df.iloc[i:i+num_rows, :] for i in split_indices]
    # Get the folder path and base name of the original Excel file
    folder_path = os.path.dirname(file_path)
    file_base_name = os.path.splitext(os.path.basename(file_path))[0]
    # Save the smaller DataFrames as Excel files in the same folder as the original file
    for i, df in enumerate(df_list):
        df.to_excel(f"{folder_path}/{file_base_name}-{i+1}.xlsx", index=False)

    completion_label.config(text="Successfully Completed!!!")


# Create the main window
window = tk.Tk()
window.title("Excel File Breakdown")
window.geometry("600x400")


# File path label
file_path_label = tk.Label(window, text="No file selected")
file_path_label.pack(padx=10, pady=10)


# Create a label and text entry for the number of rows
num_rows_label = tk.Label(text="Number of rows per Excel:", font=('Arial', 11))
num_rows_entry = tk.Entry(window, width=8)
num_rows_label.pack(padx=10, pady=10)
num_rows_entry.pack(padx=10, pady=10)


breakdown_button = tk.Button(window, text="Upload and Breakdown", command=breakdown_button_clicked, bg='light green',
                             font=('Arial', 12))
breakdown_button.pack(padx=10, pady=10)


completion_label = tk.Label(window, text="", font=('Arial', 10))
completion_label.pack(padx=10, pady=10)

# Run the main loop
window.mainloop()
