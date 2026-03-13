import os
from formatter import format_report

INPUT_FOLDER = "input"

def find_report():

    files = os.listdir(INPUT_FOLDER)

    for f in files:
        if f.endswith(".xlsx"):
            return os.path.join(INPUT_FOLDER, f)

    raise Exception("No Excel file found in input folder")

def main():
    
    report_path = find_report()
    output = format_report(report_path)

if __name__ == "__main__":
    main()