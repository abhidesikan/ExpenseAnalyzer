import openpyxl

MAX_ROW = 200
MAX_COLUMN = 5


#def get_total_expense_per_month():


#def get_total_income_per_month():


def get_month_data(sheet):
    if('October 2016' == sheet.title):
        MAX_ROW = sheet.max_row
        MAX_COLUMN = sheet.max_column
        MIN_ROW = sheet.min_row
        MIN_COLUMN = sheet.min_column
        
        for i in range(1, MAX_ROW):
#            for j in range(1, MAX_COLUMN):
                if(sheet.cell(row = i, column = 5).value is not None and sheet.cell(row = i, column = 2).value is not None):
                    print sheet.cell(row = i,column = 5).value

def load_excel_workbook():
    wb = openpyxl.load_workbook('Expenses.xlsx', data_only = True)
    return wb
    for sheet in wb.worksheets: 
        get_month_data(sheet)

def display_menu_options(wb):
    menu = {}
    menu['1'] = "Find total expense for a month"
    menu['2'] = "Find total income for a month"
    menu['3'] = "Find total expenses for a year"
    menu['4'] = "Find total income for a year"
    menu['5'] = "Find average expense for a year"
    menu['6'] = "Find average income for a year"
    
    while True:
        options = menu.keys()
        options.sort()
        for entry in options:
            print entry, menu[entry]
        
        selection = raw_input("Please select : ")
        if(selection == '1'):
            month = raw_input("Please enter month followed by year (October 2016) : ")
            sheet = wb.get_sheet_by_name(month)
            get_month_data(sheet)
        else:
            print "Unknown option selected"

def run():
    wb = load_excel_workbook()
    display_menu_options(wb)

if __name__ == '__main__':
    run()
