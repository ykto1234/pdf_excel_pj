import os
import openpyxl
import win32com.client        #win32comをインポートするだけでは上手くいかないので注意！！

def copy_location_number(input_excel_path, input_sheet):
    workbook = openpyxl.load_workbook(input_excel_path)
    sheet = workbook[input_sheet]

    numbers = []
    start_row = 8
    end_row = 1000
    number_col = 2

    for i in range(start_row, end_row):
        cell_value = sheet.cell(row=i, column=number_col).value

        if cell_value is None:
            break

        if cell_value not in numbers:
            numbers.append(cell_value)

    return numbers

def paste_location_number(numbers, input_excel_path, input_sheet, work_filename):
    workbook = openpyxl.load_workbook(input_excel_path)
    sheet = workbook[input_sheet]

    for i in range(2, 2 + len(numbers)):
        sheet.cell(row=i, column=1).value = numbers[i - 2]

    workbook.save(filename = work_filename)

def excel_to_pdf(work_filename, input_sheet, output_pdf_path):
    excel = win32com.client.Dispatch("Excel.Application")
    #excel.Visible = True
    file = excel.Workbooks.Open(os.getcwd() + '/' + work_filename)
    file.WorkSheets[input_sheet].Select()
    outpath = os.getcwd() + '/' + output_pdf_path
    file.ActiveSheet.ExportAsFixedFormat(0, outpath)
    excel.DisplayAlerts = False
    excel.Quit()
