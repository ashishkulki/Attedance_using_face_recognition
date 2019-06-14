import xlsxwriter
import xlrd
from xlutils.copy import copy
from datetime import datetime, date, time
import time
import os.path
save_path ='/home/pi/abc/excel/'
filename = "Attendance" + str(datetime.now())+ ".xls"
file_name = os.path.join(save_path, filename)
file_name = str(file_name)
print file_name
x=os.path.isfile(file_name)
if x==True:
    Trb = xlrd.open_workbook(filename)
    Twb = copy(Trb)
    Tw_sheet = Twb.get_sheet(0)
    Tw_sheet.write(1,2,"ABSENT")
    Twb.save(filename)
    
    
else:
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold':1})
    center = workbook.add_format({'align':'center'})
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    time_format = workbook.add_format({'num_format': 'hh:mm:ss AM/PM'})
    worksheet.set_column(1,15)
    worksheet.set_column('B1:C1',30)
    worksheet.set_column('D1:E1',15)
    worksheet.set_column('E1:F1',15)
    
        

    worksheet.write_rich_string('A1',bold,'Sr.No.',center)
    worksheet.write_rich_string('A2',bold,'1',center)
    worksheet.write_rich_string('A3',bold,'2',center)
    worksheet.write_rich_string('A4',bold,'3',center)
    worksheet.write_rich_string('A5',bold,'4',center)
    worksheet.write_rich_string('A6',bold,'5',center)
    worksheet.write_rich_string('A7',bold,'6',center)
    worksheet.write_rich_string('A8',bold,'7',center)
    worksheet.write_rich_string('A9',bold,'8',center)
    worksheet.write_rich_string('A10',bold,'9',center)
    worksheet.write_rich_string('A11',bold,'10',center)
    worksheet.write_rich_string('B1',bold,'student name',center)
    worksheet.write_rich_string('B2',bold,'mani',center)
    worksheet.write_rich_string('B3',bold,'Abhishek',center)
    worksheet.write_rich_string('B4',bold,'neha',center)
    worksheet.write_rich_string('B5',bold,'mansi',center)
    worksheet.write_rich_string('B6',bold,'datta',center)
    worksheet.write_rich_string('B7',bold,'chandu',center)
    worksheet.write_rich_string('B8',bold,'akshay',center)
    worksheet.write_rich_string('B9',bold,'rohit',center)
    worksheet.write_rich_string('B10',bold,'dhanush',center)
    worksheet.write_rich_string('B11',bold,'yogesh',center)
    worksheet.write_rich_string('C1',bold,'Attendance status',center)
    worksheet.write_rich_string('C2',bold,'ABSENT',center)
    worksheet.write_rich_string('C3',bold,'ABSENT',center)
    worksheet.write_rich_string('C4',bold,'ABSENT',center)
    worksheet.write_rich_string('C5',bold,'ABSENT',center)
    worksheet.write_rich_string('C6',bold,'ABSENT',center)
    worksheet.write_rich_string('C7',bold,'ABSENT',center)
    worksheet.write_rich_string('C8',bold,'ABSENT',center)
    worksheet.write_rich_string('C9',bold,'ABSENT',center)
    worksheet.write_rich_string('C10',bold,'ABSENT',center)
    worksheet.write_rich_string('C11',bold,'ABSENT',center)
    
    
                
    workbook.close()
