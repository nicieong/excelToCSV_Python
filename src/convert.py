from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import sys
import pandas as pd
import os

root = Tk()
root.title("Excel to CSV File Converter")
root.geometry("1200x400")
root['background'] = "#FFFFFF"

input_excel_file_label = Label(root, text="Select your Excel file: ", width=0, font=("Helvetica"))
input_excel_file_label.grid(row=1, column=0, sticky=W, padx=3, pady=0)

input_excel_filepath_entry = Label(root, text="" , font=("Helvetica"))
input_excel_filepath_entry.grid(row=1, column=1, sticky=W, padx=3, pady=0)

target_csv_file_label = Label(root, text="Select folder to save your CSV file: ", width=0, font=("Helvetica"))
target_csv_file_label.grid(row=2, column=0, sticky=W, padx=3, pady=0)

target_csv_directory_entry = Label(root, text="" , font=("Helvetica"))
target_csv_directory_entry.grid(row=2, column=1, sticky=W, padx=3, pady=0)

csv_filename_label = Label(root, text="Enter CSV filename: ", width=0, font=("Helvetica"))
csv_filename_label.grid(row=3, column=0, sticky=W, padx=3, pady=0)

csv_filename_input = StringVar()
csv_filename_entry = Entry(root, textvariable=csv_filename_input, width=20, font=("Helvetica"))
csv_filename_entry.grid(row=3, column=1, sticky=W, padx=3, pady=0)

excel_filepath = ""
csv_filename = ""

def select_file():
    global excel_filepath
    excel_filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if excel_filepath:
        input_excel_filepath_entry.config(text="Selected File: " + excel_filepath)
    else:
        print("No file selected.")
        return

select_file_button = Button(root, text="Select File", font=("Helvetica"), command=select_file)
select_file_button.grid(row=1, column=2, sticky=W, padx=3, pady=0)

directory_path = ""

def select_directory():
    global directory_path
    directory_path = filedialog.askdirectory()
    if directory_path:
        target_csv_directory_entry.config(text="Selected Folder: " + directory_path)
    else:
        print("No folder selected.")
        return

select_directory_button = Button(root, text="Select Folder", font=("Helvetica"), command=select_directory)
select_directory_button.grid(row=2, column=2, sticky=W, padx=3, pady=0)

def convert():
    global excel_filepath, directory_path, csv_filename
    
    if not excel_filepath:
        print("No Excel file selected.")
        return

    if not directory_path:
        print("No directory selected.")
        return

    csv_filename = csv_filename_entry.get()
    if not csv_filename:
        print("Please enter a CSV filename.")
        return
    
    try:
        df = pd.read_excel(excel_filepath)
    except FileNotFoundError:
        print("File not found. Please check the file path.")
        exit()
    
    csv_full_filename = csv_filename + ".csv"
    csv_file_path = os.path.join(directory_path, csv_full_filename)
    
    try:
        df.to_csv(csv_file_path, index=False)
        print(f'CSV file saved successfully: {csv_file_path}')
        messagebox.showinfo("Success", f'CSV file saved successfully: {csv_full_filename} under {directory_path}')
    except Exception as e:
        print(f"Error occurred while saving CSV file: {e}")
        messagebox.showerror("Error", f"Error occurred while saving CSV file: {e}")

def exit():
    root.destroy()
    sys.exit(0)

convert_button = Button(root, text="Convert", font=("Helvetica", 15), command=convert) 
convert_button.place(x=5, y=150)

exit_button = Button(root, text="Exit", font=("Helvetica", 15), command=exit) 
exit_button.place(x=5, y=200) 
    
root.mainloop() 