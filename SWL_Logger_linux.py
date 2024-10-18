import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox

import openpyxl
from openpyxl import Workbook


# Function to update time
def update_time():
    current_time = datetime.now().strftime("%H:%M")
    entry_time.delete(0, tk.END)  # Clears inputbox
    entry_time.insert(0, current_time)  # Fills inputbox with current time
    root.after(60000, update_time)  # Recalls function every minute


# Function to update date
def update_date():
    current_date = datetime.now().strftime("%d-%m-%Y")
    entry_date.delete(0, tk.END)  # Clears inputbox
    entry_date.insert(0, current_date)  # Fills inputbox with current date
    root.after(60000, update_date)  # Recalls function every minute


# List for saving reports
current_session_reports = []


# Function to put in data and save in Excel
def save_swl_report():
    callsign_1 = entry_callsign_1.get()
    name_1 = entry_name_1.get()
    qth_locator_1 = entry_qth_locator_1.get()
    callsign_2 = entry_callsign_2.get()
    name_2 = entry_name_2.get()
    qth_locator_2 = entry_qth_locator_2.get()
    date = entry_date.get()
    time = entry_time.get()
    frequency = entry_frequency.get()
    mode = entry_mode.get()
    readability = entry_readability.get()
    signal_strength = entry_signalstrength.get()
    tone = entry_tone.get()
    details = entry_details.get()

    # Try to save reports to file in Excel
    try:
        workbook = openpyxl.load_workbook("swl_reports.xlsx")
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        # Add header
        sheet.append(
            ["Callsign 1", "Name 1", "QTH Locator 1", "Callsign 2", "Name 2", "QTH Locator 2", "Date", "Time (UTC)",
             "Freq. (MHz)", "Mode", "Readability", "Signal strength", "Tone", "Details"])

    # Add new report
    sheet.append([callsign_1, name_1, qth_locator_1, callsign_2, name_2, qth_locator_2, date, time, frequency, mode,
                  readability, signal_strength, tone, details])

    # Save file
    workbook.save("swl_reports.xlsx")

    # Add report
    current_session_reports.append(
        [callsign_1, name_1, qth_locator_1, callsign_2, name_2, qth_locator_2, date, time, frequency, mode, readability,
         signal_strength, tone, details])

    # Empty inputbox for new report
    entry_callsign_1.delete(0, tk.END)
    entry_name_1.delete(0, tk.END)
    entry_qth_locator_1.delete(0, tk.END)
    entry_callsign_2.delete(0, tk.END)
    entry_name_2.delete(0, tk.END)
    entry_qth_locator_2.delete(0, tk.END)
    entry_date.delete(0, tk.END)
    entry_time.delete(0, tk.END)
    entry_frequency.delete(0, tk.END)
    entry_mode.set("")  # Reset dropdown
    entry_readability.delete(0, tk.END)
    entry_signalstrength.delete(0, tk.END)
    entry_tone.delete(0, tk.END)
    entry_details.delete(0, tk.END)

    # Add date again
    entry_date.insert(0, datetime.now().strftime("%d-%m-%Y"))

    # Add time again
    entry_time.insert(0, datetime.now().strftime("%H:%M"))

    # Update report lists
    update_report_list()


# Function to open Excel file
def open_excel_file():
    try:
        os.system("xdg-open swl_reports.xlsx")  # Works on linux
    except Exception as e:
        messagebox.showerror("Error", f"Couldn't open file: {e}")


# Function to update list
def update_report_list():
    # Empty list
    report_list.delete(0, tk.END)
    # Add current reports to list
    for report in current_session_reports:
        report_list.insert(tk.END, report)  # Add each report to the list


# Function to go to next field
def next_field(event):
    event.widget.tk_focusNext().focus()
    return "break"


# Function for "About" window
def show_about():
    about_window = tk.Toplevel(root)
    about_window.title("About this tool")

    # Information about the maker
    about_info = """SWL Logging Tool
    Designed bij Steven Duyck. ONL13316.
    Membership UBA since August 1st 2024.
    This program logs SWL reports into an Excel-file.

    Special thanks to ChatGPT and OpenAI for support."""

    tk.Label(about_window, text=about_info, padx=20, pady=20).pack()

    close_button = tk.Button(about_window, text="Close", command=about_window.destroy)
    close_button.pack(pady=10)


# Interface with Tkinter
root = tk.Tk()
root.title("SWL Logging Tool")

# Labels and input boxes
tk.Label(root, text="Callsign 1:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_callsign_1 = tk.Entry(root)
entry_callsign_1.grid(row=0, column=1, padx=5, pady=5)
entry_callsign_1.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Name 1:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_name_1 = tk.Entry(root)
entry_name_1.grid(row=1, column=1, padx=5, pady=5)
entry_name_1.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="QTH Locator 1:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_qth_locator_1 = tk.Entry(root)
entry_qth_locator_1.grid(row=2, column=1, padx=5, pady=5)
entry_qth_locator_1.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Callsign 2:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_callsign_2 = tk.Entry(root)
entry_callsign_2.grid(row=3, column=1, padx=5, pady=5)
entry_callsign_2.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Name 2:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_name_2 = tk.Entry(root)
entry_name_2.grid(row=4, column=1, padx=5, pady=5)
entry_name_2.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="QTH Locator 2:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_qth_locator_2 = tk.Entry(root)
entry_qth_locator_2.grid(row=5, column=1, padx=5, pady=5)
entry_qth_locator_2.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Date (DD-MM-YYYY):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
entry_date = tk.Entry(root)
entry_date.grid(row=6, column=1, padx=5, pady=5)
entry_date.insert(0, datetime.now().strftime("%d-%m-%Y"))
entry_date.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Time (UTC, HH:MM):").grid(row=5, column=0, padx=5, pady=5, sticky="e")
entry_time = tk.Entry(root)
entry_time.grid(row=7, column=1, padx=5, pady=5)
entry_time.insert(0, datetime.now().strftime("%H:%M"))
entry_time.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Frequency (MHz):").grid(row=6, column=0, padx=5, pady=5, sticky="e")
entry_frequency = tk.Entry(root)
entry_frequency.grid(row=8, column=1, padx=5, pady=5)
entry_frequency.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Mode:").grid(row=7, column=0, padx=5, pady=5, sticky="e")
entry_mode = tk.StringVar(root)
mode_menu = tk.OptionMenu(root, entry_mode, "SSB", "CW", "AM", "FM", "Digital")  # Different modi
mode_menu.grid(row=9, column=1, padx=5, pady=5)

tk.Label(root, text="Readability:").grid(row=8, column=0, padx=5, pady=5, sticky="e")
entry_readability = tk.Entry(root)
entry_readability.grid(row=10, column=1, padx=5, pady=5)
entry_readability.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Signal strength:").grid(row=9, column=0, padx=5, pady=5, sticky="e")
entry_signalstrength = tk.Entry(root)
entry_signalstrength.grid(row=11, column=1, padx=5, pady=5)
entry_signalstrength.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Tone:").grid(row=10, column=0, padx=5, pady=5, sticky="e")
entry_tone = tk.Entry(root)
entry_tone.grid(row=12, column=1, padx=5, pady=5)
entry_tone.bind("<Return>", next_field)  # Connect Enter

tk.Label(root, text="Details:").grid(row=11, column=0, padx=5, pady=5, sticky="e")
entry_details = tk.Entry(root)
entry_details.grid(row=13, column=1, padx=5, pady=5)
entry_details.bind("<Return>", next_field)  # Connect Enter

# Listbox for reports
report_list = tk.Listbox(root, height=10, width=70)
report_list.grid(row=14, column=0, columnspan=2, padx=10, pady=5)

# Save button
save_button = tk.Button(root, text="Save", command=save_swl_report)
save_button.grid(row=15, column=0, padx=5, pady=5, sticky="ew")

# Open button
open_button = tk.Button(root, text="Open logfile", command=open_excel_file)
open_button.grid(row=15, column=1, padx=5, pady=5, sticky="ew")

# About button
about_button = tk.Button(root, text="About", command=show_about)
about_button.grid(row=16, column=0, padx=5, pady=5, sticky="ew")

# Exit button
exit_button = tk.Button(root, text="Exit", command=root.quit)
exit_button.grid(row=16, column=1, padx=5, pady=5, sticky="ew")


# Center buttons
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Start the interface
root.mainloop()
