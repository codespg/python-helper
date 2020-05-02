import openpyxl
import os
import csv


def xlsx_to_csv_by_xlrd(excel_file, csv_file):
    csv_file_open = open(csv_file, 'w', encoding="utf-8")
    csv_writer = csv.writer(csv_file_open, quoting=csv.QUOTE_ALL)

    workbook = openpyxl.load_workbook(excel_file)
    for worksheet in workbook.worksheets:
        # Check if Visible sheet of Excel
        if not worksheet.sheet_state == 'hidden':
            for row in worksheet.rows:
                csv_writer.writerow([cell.value for cell in row])
    csv_file_open.close()
    return csv_file


if __name__ == '__main__':
    excel_file = input("Please enter excel file to convert:")
    csv_file = os.path.splitext(excel_file)[0] + '.csv'

    result = xlsx_to_csv_by_xlrd(excel_file, csv_file)
    print("CSV Converted Successfully :", result)
