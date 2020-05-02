import xlrd
import os
import csv


def xlsx_to_csv_by_xlrd(excel_file, csv_file):
    csv_file_open = open(csv_file, 'w', encoding="utf-8")
    csv_writer = csv.writer(csv_file_open, quoting=csv.QUOTE_ALL)

    workbook = xlrd.open_workbook(excel_file)
    for sheet_name in workbook.sheet_names():
        sheet = workbook.sheet_by_name(sheet_name)
        # Check if Visible sheet of Excel
        if sheet.visibility == 0:
            for rownum in range(sheet.nrows):
                csv_writer.writerow(sheet.row_values(rownum))
    csv_file_open.close()
    return csv_file


if __name__ == '__main__':
    excel_file = input("Please enter excel file to convert:")
    csv_file = os.path.splitext(excel_file)[0] + '.csv'

    result = xlsx_to_csv_by_xlrd(excel_file, csv_file)
    print("CSV Converted Successfully :", result)
