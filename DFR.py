import os

 # Here's the change: Specify the path to the mobile Python interpreter
mobile_python_path = os.path.join("Documents", "ELYSIUM Python", "python") # Assuming 'python' is the mobile Python interpreter executable

import openpyxl
import tkinter as tk
from tkinter import filedialog, ttk
import time
import csv
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from PyQt5.QtWidgets import QApplication, QWidget
import datetime
import re

# Initialize the GUI application
root = tk.Tk()
root.title("CSV File Editor")
window_width = 400
window_height = 290
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = int((screen_width - window_width) / 2)
y_coordinate = int((screen_height - window_height) / 2)
root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

# Create a ttk.Style object
style = ttk.Style()

# Initially define the theme
current_theme = "light"

# Define color schemes for dark and light modes, including ttk styles
color_schemes = {
    "light": {
        "bg": "#FFFFFF",
        "fg": "#000000",
        "button_bg": "#4CAF50",
        "button_fg": "white",
        "ttk_theme": "default"
    },
    "dark": {
        "bg": "#333333",
        "fg": "#CCCCCC",
        "button_bg": "#555555",
        "button_fg": "white",
        "ttk_theme": "clam"
    }
}

progress_var = tk.DoubleVar()

# Function to toggle and apply themes
def toggle_theme():
    global current_theme
    current_theme = "dark" if current_theme == "light" else "light"
    apply_theme()

def apply_theme():
    theme = color_schemes[current_theme]
    root.configure(bg=theme["bg"])
    for widget in root.winfo_children():
        if isinstance(widget, tk.Button) or isinstance(widget, tk.Label):
            widget.configure(bg=theme["bg"], fg=theme["fg"])
        if isinstance(widget, tk.Text):
            widget.configure(bg=theme["bg"], fg=theme["fg"])
    # Configure button specific colors
    select_button.configure(bg=color_schemes[current_theme]["button_bg"], fg=color_schemes[current_theme]["button_fg"])
    select_template_button.configure(bg=color_schemes[current_theme]["button_bg"], fg=color_schemes[current_theme]["button_fg"])
    edit_button.configure(bg=color_schemes[current_theme]["button_bg"], fg=color_schemes[current_theme]["button_fg"])
    toggle_theme_button.configure(bg=color_schemes[current_theme]["button_bg"], fg=color_schemes[current_theme]["button_fg"])
    # Configure progress bar style
    if current_theme == "dark":
        style_name = "Dark.Horizontal.TProgressbar"
        style.configure(style_name, background="#555555", troughcolor="#333333")
        progress_bar.configure(style=style_name)
    else:
        style_name = "Light.Horizontal.TProgressbar"
        style.configure(style_name, background="#FFFFFF", troughcolor="#E0E0E0")
        progress_bar.configure(style=style_name)

def update_progress(progress_var, value, max_value):
    progress = (value / max_value) * 100
    progress_var.set(progress)
    root.update_idletasks()

def sort_sheet(worksheet):
    def custom_sort(row):
        sort_columns = [8, 6, 15, 16, 20, 21]
        sort_values = [row[col - 1] for col in sort_columns]
        return sort_values
    rows_to_sort = list(worksheet.iter_rows(min_row=2, values_only=True))
    new_rows = []
    prev_col_f_value = None
    for row_data in sorted(rows_to_sort, key=custom_sort):
        col_f_value = row_data[5]
        if prev_col_f_value is not None and col_f_value != prev_col_f_value:
            # Insert a blank row with 'a' in Column F
            new_separator_row = [''] * len(row_data)
            new_separator_row[5] = 'a'
            new_separator_row[46] = None
            new_rows.append(new_separator_row)
            for col_index in range(1, len(row_data) + 1):
                if col_index != 47:
                    worksheet.cell(row=len(new_rows) + 1, column=col_index).fill = openpyxl.styles.PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        new_rows.append(row_data)
        prev_col_f_value = col_f_value

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        for cell in row:
            cell.value = None
    for row_index, row_data in enumerate(new_rows, start=2):
        for col_index, cell_value in enumerate(row_data, start=1):
            cell = worksheet.cell(row=row_index, column=col_index, value=cell_value)
    for col_index in range(39, 48):
        col_letter = openpyxl.utils.get_column_letter(col_index)
        for row_index in range(2, worksheet.max_row + 1):
            cell = worksheet[f"{col_letter}{row_index}"]
            cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thick'))
            if row_index == worksheet.max_row:
                cell.border += openpyxl.styles.Border(bottom=openpyxl.styles.Side(style='thick'))
        for row_index in range(2, worksheet.max_row + 1):
            cell = worksheet[f"{col_letter}{row_index}"]
            cell.border += openpyxl.styles.Border(right=openpyxl.styles.Side(style='thick'))
    for col_index in range(40, 46):
        col_letter = openpyxl.utils.get_column_letter(col_index)
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=col_index, max_col=col_index):
            for cell in row:
                cell.number_format = '0%'
def center_all_cells(worksheet):
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
def hide_columns(worksheet, columns_to_hide):
    for col_index in columns_to_hide:
        col_letter = openpyxl.utils.get_column_letter(col_index)
        worksheet.column_dimensions[col_letter].hidden = True
make_mapping = {
    "AMC": "AMC",
    "ACUR": "Acura",
    "AUDI": "Audi",
    "BUIC": "Buick",
    "CADI": "Cadillac",
    "CHEV": "Chevrolet",
    "CHRY": "Chrysler",
    "DODG": "Dodge",
    "FORD": "Ford",
    "GMC": "GMC",
    "HOND": "Honda",
    "HYUN": "Hyundai",
    "INFI": "Infiniti",
    "JAG": "Jaguar",
    "JEEP": "Jeep",
    "KIA" : "Kia",
    "LEXU": "Lexus",
    "LINC": "Lincoln",
    "LR": "Land Rover",
    "MAZD": "Mazda",
    "MERC": "Mercury",
    "MITS": "Mitsubishi",
    "NISS": "Nissan",
    "OLDS": "Oldsmobile",
    "PLYM": "Plymouth",
    "PONT": "Pontiac",
    "RAM": "Ram",
    "SATN": "Saturn",
    "SUBA": "Subaru",
    "TOYO": "Toyota",
    "VW": "Volkswagen"
}
def edit_file(file_path, template_path=None):
    try:
        start_time = time.time()
        print("Reading CSV file...")
        update_progress(progress_var, 1, 5)
        with open(file_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            data = list(reader)
        # Change names in the data based on the mapping
        for row in data:
            for i in range(len(row)):
                current_input = row[i]
                if current_input and current_input in make_mapping:
                    row[i] = make_mapping[current_input]

        print("CSV file read successfully.")
        for row in data:
            if len(row) > 6:
                del row[6]
        excel_file_path = os.path.join(os.path.dirname(file_path), "DFR Month Editing in Progress.xlsx")
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        for row_data in data:
            worksheet.append(row_data)
        worksheet.freeze_panes = 'A2'
        update_progress(progress_var, 2, 5)
        for row_index in range(worksheet.max_row, 1, -1):
            if worksheet.cell(row=row_index, column=16).value == '':
                worksheet.delete_rows(row_index)
        workbook.save(excel_file_path)
        print("Filtered sheet created.")
        hide_columns(workbook.active, [1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 21, 30, 31, 32, 33, 34, 35, 36, 37, 38])
        filtered_sheet = workbook.create_sheet(title='False Negatives')
        max_columns = max(len(row) for row in data)
        for row_data in data:
            row_data.extend([''] * (max_columns - len(row_data)))
            if not row_data or row_data[15] == '':
                filtered_sheet.append(row_data)
        workbook.save(excel_file_path)
        print("Template processing started.")
        update_progress(progress_var, 3, 5)
        if template_path:
            template_wb = openpyxl.load_workbook(template_path)
            template_ws = template_wb.active
            header_row = [cell.value if isinstance(cell, openpyxl.cell.cell.Cell) else cell for cell in template_ws[1]]
            edited_wb = openpyxl.load_workbook(excel_file_path)
            edited_ws = edited_wb.active
            for col_num, value in enumerate(header_row, 1):
                cell = edited_ws.cell(row=1, column=col_num, value=value)
                cell.font = openpyxl.styles.Font(bold=True)
                cell.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side(style='medium'))
            edited_ws.auto_filter.ref = edited_ws.dimensions
            sort_sheet(edited_ws)
            center_all_cells(edited_ws)
            current_month = datetime.datetime.now().strftime("%B")
            new_file_name = os.path.join(os.path.dirname(file_path), f"DFR {current_month}.xlsx")
            edited_wb.save(new_file_name)
            print(f"Saved edited file as {new_file_name}")
        update_progress(progress_var, 4, 5)
        os.remove(file_path)
        os.remove(excel_file_path)
        print("Deleted original CSV and edited Excel files.")
    except Exception as e:
        print("An error occurred during editing:")
        print(e)
    finally:
        update_progress(progress_var, 5, 5)
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        file_display_text.config(state=tk.NORMAL)
        file_display_text.delete(1.0, tk.END)
        file_display_text.insert(tk.END, file_path)
        file_display_text.config(state=tk.DISABLED)
def select_template():
    template_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if template_path:
        template_display_text.config(state=tk.NORMAL)
        template_display_text.delete(1.0, tk.END)
        template_display_text.insert(tk.END, template_path)
        template_display_text.config(state=tk.DISABLED)
def edit_file_wrapper():
    file_path = file_display_text.get("1.0", "end-1c")
    template_path = template_display_text.get("1.0", "end-1c")
    if file_path:
        edit_file(file_path, template_path)

# GUI component definitions with corrected commands for functionality
title_label = tk.Label(root, text="Attach the CSV File needing Editing", font=("Arial", 14))
title_label.grid(row=0, column=0, columnspan=2, pady=10)

select_button = tk.Button(root, text="Select File", command=select_file, padx=10, pady=5, bg="#4CAF50", fg="white")
select_button.grid(row=1, column=0, padx=10, pady=5)

file_display_text = tk.Text(root, height=1, width=20, wrap=tk.WORD, state=tk.DISABLED)
file_display_text.grid(row=1, column=1, padx=10, pady=5)

file_scrollbar = tk.Scrollbar(root, command=file_display_text.yview)
file_scrollbar.grid(row=1, column=2, sticky='nsew')
file_display_text['yscrollcommand'] = file_scrollbar.set

select_template_button = tk.Button(root, text="Select Template for Header", command=select_template, padx=10, pady=5, bg="#4CAF50", fg="white")
select_template_button.grid(row=2, column=0, padx=10, pady=5)

template_display_text = tk.Text(root, height=1, width=20, wrap=tk.WORD, state=tk.DISABLED)
template_display_text.grid(row=2, column=1, padx=10, pady=5)

template_scrollbar = tk.Scrollbar(root, command=template_display_text.yview)
template_scrollbar.grid(row=2, column=2, sticky='nsew')
template_display_text['yscrollcommand'] = template_scrollbar.set

edit_button = tk.Button(root, text="Edit File", command=edit_file_wrapper, padx=10, pady=5, bg="#007BFF", fg="white")
edit_button.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate", variable=progress_var)
progress_bar.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Add a toggle theme button
toggle_theme_button = tk.Button(root, text="Dark/Light", command=toggle_theme, padx=10, pady=5)
toggle_theme_button.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

# Apply initial theme
apply_theme()

root.mainloop()

