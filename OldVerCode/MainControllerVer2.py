#New file on maincontroller Version 1.2
#Dec 10 2020


import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side

import os
from os import listdir
from os.path import isfile, join

# Nhap Pyexcel
import pyexcel as p

# Nhap Win32 de mo file sua loi
import win32com.client as win32






#Lay thong tin filenames
input_path = 'input'
input_files = [f for f in listdir(input_path) if isfile(join(input_path,f))]


input_cooked_files = []
for item in input_files:
    raw = input_path + '\\' + item
    input_cooked_files.append(raw)




#Kiem tra neu file da duoc chuyen sang xlsx
check_list = []
for item in input_cooked_files:
    if item.find('.xlsx') != -1:
        check_list.append(item)

#print(check_list)

output_temp = 'Not_touch_temp'
output_cooked_path = []
for item in input_files:
    if item.find('.xlsx') == -1:
        raw = output_temp + '\\' + item + 'x'
        output_cooked_path.append(raw)


# Chuyen toan bo file sang xlsx
for id in range(0,len(output_cooked_path)):
    #Lệnh lấy path nguồn từ disk cho máy
    cd = os.path.dirname(os.path.abspath(__file__))
    if (output_cooked_path[id] not in check_list) and (input_cooked_files[id].find('.xlsx') == -1):

        try:
            f = os.path.join(cd, input_cooked_files[id])
            xl = win32.gencache.EnsureDispatch('Excel.Application')
            wb = xl.Workbooks.Open(f)
            print(f)
            xl.ActiveWorkbook.SaveAs(os.path.join(cd, output_cooked_path[id]), FileFormat=56)
            wb.Close(True)

        except Exception as e:
            print(e)

        finally:
            wb = None
            xl = None


        #excel_app = win32.Dispatch('Excel.Application')
        #wb = excel_app.Workbooks.open(input_cooked_files[id])
        #excel_app.DisplayAlerts = False
        #wb.Save()
        #excel_app.quit()
        #p.save_book_as(file_name=input_cooked_files[id],dest_file_name=output_cooked_path[id])


#print(check_list)
#print(output_cooked_path)
#print(input_cooked_files)
#print(input_files)


