###################################################################
# Author: Caleb Sellinger
# Date: 2026-06-11
#
# Description:
# Automation for export file cleanup
# Notes:
#
# Known Issues:
#
###################################################################

import sys
import pandas as pd
from tkinter import Tk, filedialog, messagebox, StringVar
from tkinter import *
from pathlib import Path
import os
import subprocess

def gui_read_file():
    """Simple GUI to select the export file for cleanup."""

    root = Tk()
    root.title("Export File Cleanup")
    root.update_idletasks()
    # get screen width and height
    screen_width = root.winfo_screenwidth() // 2
    screen_height = root.winfo_screenheight() // 2
    x = (root.winfo_screenwidth() // 2) - (screen_width // 2)
    y = (root.winfo_screenheight() // 2) - (screen_height // 2)
    # set the size of the window to half scale and position of the window to the center of the screen
    root.geometry(f"{screen_width// 2}x{screen_height // 2}+{x}+{y}")

    # make grid expand so widgets grow with the window
    for i in range(4):
        root.columnconfigure(i, weight=1)
    for i in range(4):
        root.rowconfigure(i, weight=1)

    file_input1 = StringVar(value="")

    def browse_file():
        """Open a file dialog to select the export file."""
        file_path = filedialog.askopenfilename(
            title="Select Export File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
            initialdir=str(Path.home() / "Downloads")
        )
        if file_path:
            file_input1.set(file_path)
            input_label.config(text=f"{file_path}")

    def on_run():
        if not file_input1.get():
            messagebox.showerror("Error", "Please select an input file.")
            return
        root.quit()

    def on_cancel():
        file_input1.set("")
        root.destroy()
        sys.exit()

    # Button for Input File
    input_label = Label(root, text="No input file selected", wraplength=700, anchor="w", justify="left")
    input_label.grid(row=0, column=0, columnspan=3, sticky="we", padx=8, pady=8)
    Button(root, text="Choose Input File", command=browse_file).grid(row=0, column=3, sticky="e", padx=8, pady=8)

    # Center the Run button across the dialog
    Button(root, text="         Run         ", command=on_run).grid(row=4, column=0, columnspan=4, pady=12)
    # Place Cancel below/right
    Button(root, text="         Cancel       ", command=on_cancel).grid(row=4, column=3, pady=8, padx=8, sticky="e")
    # Handles if exit button is clicked instead of cancel button
    root.protocol("WM_DELETE_WINDOW", on_cancel)

    root.mainloop()

    if not file_input1.get():
        sys.exit(0)

    return read_file(file_input1.get())

def read_file(input_file: StringVar) -> pd.DataFrame:
    """
    Read the selected export file and perform cleanup operations.
    """
    try:
            if os.path.isabs(input_file):
                file_path = input_file
            else:
                file_path = os.path.join(os.path.dirname(__file__), input_file)
                print(f"Resolved file path: {file_path}")
            
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)

            return df

    except Exception as e:
        messagebox.showerror("Error reading file", f"An error occurred while processing the file: {e}")
        sys.exit(0)

def cleanup_data(df: pd.DataFrame):
    """
    Perform data cleanup operations on the DataFrame.
    """
    # Always output to download folder
    with pd.ExcelWriter(Path.home() / 'Downloads' / 'cleaned_data.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Sheet 1", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Sheet 1"]

        #freeze the first row
        worksheet.freeze_panes(1, 0)

    # Opens the output file in the default application (Excel)
    # os.startfile(Path.home() / 'Downloads' / 'cleaned_data.xlsx')

    # Opens the output file in explorer and selects it
    subprocess.Popen(f'explorer /select,"{Path.home() / "Downloads" / "cleaned_data.xlsx"}"', shell=True)

def main():
    df = gui_read_file()
    cleanup_data(df)

if __name__ == "__main__":
    main()