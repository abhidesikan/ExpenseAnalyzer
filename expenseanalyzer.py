import openpyxl

MAX_ROW = 200
MAX_COLUMN = 5

def get_month_data(sheet):
    if('October 2016' == sheet.title):
        MAX_ROW = sheet.max_row
        MAX_COLUMN = sheet.max_column
        print MAX_ROW
        print MAX_COLUMN
        for i in range(1, MAX_ROW):
            for j in range(1, MAX_COLUMN):
                print sheet.cell(row=i,column=j).value

def load_excel_workbook():
    wb = openpyxl.load_workbook('Expenses.xlsx')
    for sheet in wb.worksheets: 
        get_month_data(sheet)

def run():
    load_excel_workbook()

if __name__ == '__main__':
    run()
