import openpyxl

def get_row_count(filename, sheetname):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.get_sheet_by_name(sheetname)
    return sheet.max_row

def get_column_count(filename, sheetname):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.get_sheet_by_name(sheetname)
    return sheet.max_column

def read_data(filename, sheetname, rownum, columnnum):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.get_sheet_by_name(sheetname)
    return sheet.cell(row=rownum, column=columnnum).value

def write_data(filename, sheetname, rownum, columnnum, input_data):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.get_sheet_by_name(sheetname)
    sheet.cell(row=rownum, column=columnnum).value=input_data
    workbook.save(filename)