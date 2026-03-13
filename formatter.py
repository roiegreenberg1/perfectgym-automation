import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side

def format_report(report_path, config):

    data_frame = pd.read_excel(report_path)

    # Filters out the columns not required and then custom sorts the remaining rows
    data_frame = data_frame[config["columns_needed"]]
    data_frame = data_frame.sort_values(
        by=["Time", "Class", "Zone"],
        ascending=[True, True, True]
    )

    # Creates a new .xlsx file and adds data from the data frame
    output_folder = config["output_folder"]
    os.makedirs(output_folder, exist_ok = True)
    output_file = os.path.join(output_folder, config["output_filename"])
    data_frame.to_excel(output_file, index = False)

    workbook = load_workbook(output_file)
    worksheet = workbook.active
    apply_styling(worksheet, config)

    # Landscape printing
    worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE

    workbook.save(output_file)

def apply_styling(worksheet, config):

    column_widths = {
        "A": 15,  # Class
        "B": 12,  # Day
        "C": 12,  # Time
        "D": 10,  # Zone
        "E": 20,  # Class Trainer
        "F": 15,  # Member ID
        "G": 15,  # First Name
        "H": 15,  # Last Name
    }

    # Changes the width of each column to match the correct sizing of its header
    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    # Applies a bold styling to the cells in the first row (header row)
    for cell in worksheet[1]:
        cell.font = Font(bold = True)

    # Defines the border styling
    thin = Side(border_style = "thin", color = "000000")
    border = Border(top = thin, left = thin, right = thin, bottom = thin)

    # Applies a thin border to all cells in the worksheet
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = border

    # Finds at which position the "Class Trainer" column is found
    header_row = next(worksheet.iter_rows(min_row = 1, max_row = 1))
    column_index = next((i for i, cell in enumerate(header_row) if cell.value == "Class Trainer"), None)

    if column_index is None:
        raise ValueError("Column 'Class Trainer' not found in header row")

    # Fills in the rows with alternating colours, switching when the trainer changes
    colours = config["trainer_colours"]
    colour_index = 0
    prev_trainer = None

    for row in worksheet.iter_rows(min_row = 2):
        trainer_curr = row[column_index].value

        if trainer_curr != prev_trainer:
            colour_index = 1 - colour_index
            prev_trainer = trainer_curr

        for cell in row:
            cell.fill = PatternFill(patternType = "solid", fgColor = colours[colour_index])
