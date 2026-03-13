import os
import json
from formatter import format_report

with open("config.json") as f:
    config = json.load(f)

def find_report():
    files = os.listdir(config["input_folder"])

    # Looks for all excel files in /input and returns the file path 
    for f in files:
        if f.endswith(".xlsx"):
            return os.path.join(config["input_folder"], f)

    raise Exception("No Excel file found in input folder")

def main():
    report_path = find_report()
    format_report(report_path, config)

if __name__ == "__main__":
    main()
