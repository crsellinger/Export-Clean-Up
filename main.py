###################################################################
# Author: Caleb Sellinger
# Date: 2026-06-12
#
# Version: 1.0
# Description:
# Automation for export file cleanup
#
# Notes:
# Data cleanup for exported files from Raiser's Edge.
# The script performs the following operations:
# 1. Removes records where the primary is deceased.
# 2. Cleans up secondary contact information if deceased or ghost.
# 3. Freezes and bolds the header row.
# 4. Removes records that are not US-based.
# 5. Formats giving columns as currency.
#
# Known Issues:
#
###################################################################

import operator
import sys
import pandas as pd
from tkinter import Tk, filedialog, messagebox, StringVar, Button, Label, Entry
from pathlib import Path
import os
import subprocess


def read_file(input_file: StringVar) -> pd.DataFrame:
    """
    Read the selected export file. Supports both CSV and Excel formats.
    The filename is stored as an attribute of the DataFrame for later use when saving the cleaned file.
    """
    try:
        if os.path.isabs(input_file):
            file_path = input_file
        else:
            file_path = os.path.join(os.path.dirname(__file__), input_file)

        if file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path)
            df.attrs["filename"] = os.path.basename(file_path)
        else:
            df = pd.read_csv(file_path)
            df.attrs["filename"] = (
                os.path.basename(file_path).rsplit(".", 1)[0] + "_cleaned.xlsx"
            )

        return df

    except Exception as e:
        messagebox.showerror("Error reading file\n", f"An error reading the file:\n{e}")
        sys.exit(0)


def cleanup_data(df: pd.DataFrame, currency: StringVar) -> None:
    """
    Perform data cleanup operations on the DataFrame.
    """
    try:
        for index, row in df.iterrows():
            if df.get("Primary Deceased") is not None:
                if str(row["Primary Deceased"]) == "True":
                    df.drop(index, inplace=True)
            if df.get("Primary Address Country") is not None:
                if not operator.xor(
                    str(row["Primary Address Country"]) == "US",
                    str(row["Primary Address Country"]) == "USA",
                ):
                    df.drop(index, inplace=True)
            if df.get("Secondary Deceased") is not None:
                if str(row["Secondary Deceased"]) == "True":
                    # If the secondary is deceased, we want to keep the record but clear out the secondary fields
                    if df.get("Secondary Nickname") is not None:
                        df.at[index, "Secondary Nickname"] = ""
                    if df.get("Secondary Last Name") is not None:
                        df.at[index, "Secondary Last Name"] = ""
                    if df.get("Secondary Title") is not None:
                        df.at[index, "Secondary Title"] = ""
                    if df.get("Secondary Suffix") is not None:
                        df.at[index, "Secondary Suffix"] = ""
                    if df.get("Secondary Email Address") is not None:
                        df.at[index, "Secondary Email Address"] = ""
                    if df.get("Secondary Individual Id") is not None:
                        df.at[index, "Secondary Individual Id"] = ""
                    if df.get("Secondary First Name") is not None:
                        df.at[index, "Secondary First Name"] = ""
                    if df.get("Secondary Middle Name") is not None:
                        df.at[index, "Secondary Middle Name"] = ""
                    if df.get("Secondary Suffix") is not None:
                        df.at[index, "Secondary Suffix"] = ""
                    if df.get("Secondary Pre-Marriage Name") is not None:
                        df.at[index, "Secondary Pre-Marriage Name"] = ""
                    if df.get("Secondary Phone Number") is not None:
                        df.at[index, "Secondary Phone Number"] = ""
                    df.at[index, "Secondary Deceased"] = ""
            if df.get("Secondary Title") is not None:
                if str(row["Secondary Title"]) == "":
                    # If the secondary is ghost (denoted with no title), we want to keep the record but clear out the secondary fields
                    if df.get("Secondary Nickname") is not None:
                        df.at[index, "Secondary Nickname"] = ""
                    if df.get("Secondary Last Name") is not None:
                        df.at[index, "Secondary Last Name"] = ""
                    if df.get("Secondary Title") is not None:
                        df.at[index, "Secondary Title"] = ""
                    if df.get("Secondary Suffix") is not None:
                        df.at[index, "Secondary Suffix"] = ""
                    if df.get("Secondary Email Address") is not None:
                        df.at[index, "Secondary Email Address"] = ""
                    if df.get("Secondary Individual Id") is not None:
                        df.at[index, "Secondary Individual Id"] = ""
                    if df.get("Secondary First Name") is not None:
                        df.at[index, "Secondary First Name"] = ""
                    if df.get("Secondary Middle Name") is not None:
                        df.at[index, "Secondary Middle Name"] = ""
                    if df.get("Secondary Suffix") is not None:
                        df.at[index, "Secondary Suffix"] = ""
                    if df.get("Secondary Pre-Marriage Name") is not None:
                        df.at[index, "Secondary Pre-Marriage Name"] = ""
                    if df.get("Secondary Phone Number") is not None:
                        df.at[index, "Secondary Phone Number"] = ""
    except Exception as e:
        messagebox.showerror(
            "Error processing data\n",
            f"An error occurred while processing the data:\n{e}",
        )
        sys.exit(0)

    try:
        # Always output to download folder
        with pd.ExcelWriter(
            Path.home() / "Downloads" / df.attrs["filename"], engine="xlsxwriter"
        ) as writer:
            df.to_excel(writer, sheet_name="Sheet 1", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Sheet 1"]
            currency_format = workbook.add_format({"num_format": "$#,##0.00"})

            # Freeze the first row
            worksheet.freeze_panes(1, 0)

            # Format giving columns as currency, Col P as default
            worksheet.set_column(currency.get(), None, currency_format)

            # No border and bold headers
            format = workbook.add_format({"bold": True, "border": 0})
            for col, value in enumerate(df.columns.values):
                worksheet.write(0, col, value, format)

            # Resize columns to fit content
            worksheet.autofit()

    except Exception as e:
        messagebox.showerror(
            "Error writing file\n",
            f"An error occurred while saving the output file:\n{e}",
        )
        sys.exit(0)

    # Opens the output file in the default application (Excel)
    # os.startfile(Path.home() / 'Downloads' / df.attrs['filename'])

    # Opens the output file in explorer and selects it
    subprocess.Popen(
        f'explorer /select,"{Path.home() / "Downloads" / df.attrs["filename"]}',
        shell=True,
    )


def main():

    # Root window
    root = Tk()
    root.title("Export File Cleanup")
    root.update_idletasks()
    # get screen width and height
    screen_width = root.winfo_screenwidth() // 2
    screen_height = root.winfo_screenheight() // 2
    x = (root.winfo_screenwidth() // 2) - (screen_width // 2)
    y = (root.winfo_screenheight() // 2) - (screen_height // 2)
    # set the size of the window to half scale and position of the window to the center of the screen
    root.geometry(f"{screen_width // 2}x{screen_height // 2}+{x}+{y}")

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
            initialdir=str(Path.home() / "Downloads"),
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
    input_label = Label(
        root, text="No input file selected", wraplength=700, anchor="w", justify="left"
    )
    input_label.grid(row=0, column=0, columnspan=3, sticky="we", padx=8, pady=8)
    Button(root, text="Choose Input File", command=browse_file).grid(
        row=0, column=3, sticky="e", padx=8, pady=8
    )

    # Center the Run button across the dialog
    Button(root, text="         Run         ", command=on_run).grid(
        row=4, column=0, columnspan=4, pady=12
    )
    # Place Cancel below/right
    Button(root, text="         Cancel       ", command=on_cancel).grid(
        row=4, column=3, pady=8, padx=8, sticky="e"
    )
    # Handles if exit button is clicked instead of cancel button
    root.protocol("WM_DELETE_WINDOW", on_cancel)

    # Text input for user to input currency column letter if not standard, default to P
    currency_column = StringVar(value="")
    Label(root, text="Currency Column (format A:A):").grid(
        row=1, column=0, sticky="w", padx=8, pady=8
    )
    Entry(root, textvariable=currency_column).grid(
        row=1, column=1, sticky="we", padx=8, pady=8
    )

    root.mainloop()

    if not file_input1.get():
        sys.exit(0)

    # Testing
    # df = read_file(r"C:\Users\csellinger\Downloads\Export Test File for CleanUp.csv")
    df = read_file(file_input1.get())

    cleanup_data(df, currency_column)


if __name__ == "__main__":
    main()
