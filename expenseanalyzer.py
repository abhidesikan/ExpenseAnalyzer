import openpyxl

MAX_ROW = 200
MAX_COLUMN = 5




def get_total_income_per_month(sheet):
    MAX_ROW = sheet.max_row
    MAX_COLUMN = sheet.max_column
 
    total = 0
    for i in range(2, MAX_ROW):
        if(sheet.cell(row = i, column = 4).value is not None and sheet.cell(row = i, column = 2).value is not None):
            value = sheet.cell(row = i, column = 4).value
            total = total + float(value)
    print "The total income for month " + sheet.title + " is " + str(total)
    return total

def  get_total_expense_per_month(sheet):
    MAX_ROW = sheet.max_row
    MAX_COLUMN = sheet.max_column
       
    total = 0
    for i in range(2, MAX_ROW):
        if(sheet.cell(row = i, column = 5).value is not None and sheet.cell(row = i, column = 2).value is not None):
            value = sheet.cell(row = i, column = 5).value
            total = total + float(value)
    print "The total expenses for month " + sheet.title + " is " + str(total)
    return total

def balance(income, expense):
    print "Remaining balance is " + str(income - expense)

def load_excel_workbook():
    wb = openpyxl.load_workbook('Expenses.xlsx', data_only = True)
    return wb

def display_menu_options(wb):
    menu = {}
    menu['1'] = "Display monthly statement"
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
            income_total = get_total_income_per_month(sheet)
            expense_total = get_total_expense_per_month(sheet)
            balance(income_total, expense_total)
        else:
            print "Unknown option selected"

def run():
    wb = load_excel_workbook()
    sheet = wb.get_sheet_by_name('October 2016')
#    get_total_expense_per_month(sheet)
#    get_total_income_per_month(sheet)
    display_menu_options(wb)

if __name__ == '__main__':
    run()
