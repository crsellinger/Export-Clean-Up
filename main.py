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

import pandas as pd
from tkinter import *
from pathlib import Path
import os

def gui():
    """Simple GUI to select the export file for cleanup."""

    root = Tk()
    root.title("Export File Cleanup")
    root.update_idletasks()
    # get screen width and height
    screen_width = root.winfo_screenwidth() // 2
    screen_height = root.winfo_screenheight() // 2
    x = (root.winfo_screenwidth() // 2) - (screen_width // 2)
    y = (root.winfo_screenheight() // 2) - (screen_height // 2)
    # set the position of the window to the center of the screen
    root.geometry(f"{screen_width}x{screen_height}+{x}+{y}")

    file_input1 = StringVar(value="")
    def browse_file():
        """Open a file dialog to select the export file."""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Export File",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialdir=str(Path.home() / "Downloads"))
        if file_path:
            file_input1.set(file_path)
            input_label.config(text=f"{file_path}")
    
    input_label = Label(root, text="No input file selected", wraplength=700, anchor="w", justify="left")
    input_label.grid(row=0, column=0, columnspan=3, sticky="we", padx=8, pady=8)
    Button(root, text="Choose Degree File", command=browse_file).grid(row=0, column=3, sticky="e", padx=8, pady=8)

    root.mainloop()

def main():
    gui()

if __name__ == "__main__":
    main()