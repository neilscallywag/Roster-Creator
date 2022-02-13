from openpyxl import Workbook
import datetime
from datetime import date
from openpyxl.styles import DEFAULT_FONT
from calendar import monthrange
from openpyxl import load_workbook


wb =  Workbook()
std=wb['Sheet']
wb.remove(std)

# Variables
DEFAULT_FONT.name = 'Times New Roman'
DEFAULT_FONT.size = 9
starting_month = 2
months = 6
Date_Cell = 'F1'
Days_Cell = 'F2'
for month in range(starting_month,months+1):
    wb.create_sheet(datetime.date(1900, month, 1).strftime('%b'))
    

for sheet in wb.worksheets:
    sheet[Date_Cell] = "Date"
    sheet[Days_Cell] = "Day"
    sheet.formula_attributes['G1'] = {'ca':'1'}
    num_days = monthrange(date.today().year, datetime.datetime.strptime(sheet.title, "%b").month)[1]
    sheet['G1'] = '''=DAY(_xlfn.SEQUENCE(1,{},"1-{}-{}",1))'''.format(num_days,sheet.title,date.today().year)
    sheet.formula_attributes['G1'] = {'t': 'string'}
    


wb.save('roster.xlsx')

    
