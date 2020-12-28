#New file on maincontroller Version 1.2
#Dec 15 2020



from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side

import os
from os import listdir
from os.path import isfile, join


#Nhap cac thu vien de xu ly file xls
import pandas as pd
import xlrd
import xlwt
import openpyxl

#Cac thu vien den copy file
import datetime
import shutil
import xlsxwriter





#Xu ly file cua BO

#Doc danh sach file
input_path = 'input'
input_files = [f for f in listdir(input_path) if isfile(join(input_path,f))]

all_input_files = []
for item in input_files:
    str_input = input_path + '\\' + item
    all_input_files.append(str_input)


#Kiem tra lai danh sach file nhap vao
#for item in all_input_files:
#    print(item)

#Tien hanh lay du lieu tu file tong quat de xu ly cho bo
for tenfile in all_input_files:
    if tenfile.find("Tonghop") != -1: 
        book = xlrd.open_workbook(tenfile)
        for name in book.sheet_names():
            if name.find("Kết quả") != -1:
                #Kiem tra ten
                #print(book.sheet_by_name(name))
                
                #Tien hanh doc du lieu
                df = pd.read_excel(tenfile, sheet_name = name)
                
                #Check tinh trang sheet 
                #print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
                #print(type(df))

                #Xu ly du lieu
                row = 4
                kq = True
                while kq:
                    try:
                        int(df.iloc[row,0])
                        row += 1    
                    except:
                        kq = False
                #Chot row & col de xu ly
                final_row = row - 1
                final_col = 20
                first_row = 2
                first_col = 0

                #checkinput
                #print(row)
                
                #check dữ liệu
                #col = 0
                #while col<50:
                #    try:
                #        print(df.iloc[2,col])
                #        col += 1
                #    except:
                #        print(col)
                #        break
                
                #Doc toan bo dataframe vao xu ly
                #Lay ten lop
                cut_tenfile = tenfile.split('_')
                lop = cut_tenfile[len(cut_tenfile)-1]
                lop = lop[:-4]
                #print(lop)

                #Tao file xuat
                now = str(datetime.datetime.now())[:19]
                now = now.replace(":","_")


                src_dir0="Not_touch_core\\so_nhapdiemchitiet.xls"
                src_dir="Not_touch_core\\so_nhapdiemchitiet.xlsx"
                dst_dir="output\\output_so\\so_nhapdiemchitiet_"+lop + "_" + str(now) + ".xlsx"
                #df_title = pd.read_excel(src_dir0)
                #df_title = df_title[:7]
                #shutil.copy(src_dir,dst_dir)

                #Xoa truoc 2 dong dau cua dataframe
                df = df.iloc[2:]
                #print(df_title)
                #df_final = pd.concat([df_title,df])
                #print(df_final)
                writer = pd.ExcelWriter(dst_dir,engine = "openpyxl", mode = 'w')
                df.to_excel(writer,sheet_name = 'Sheet1', startrow=7, startcol=0, header=False, index=False)
                writer.save()
                
                #Xu ly tieu de



#Mau xu ly
#book = xlrd.open_workbook("input\Sodiem_Tonghop_10A1.xls")
#print("The number of worksheets is {0}".format(book.nsheets))
#print("Worksheet name(s): {0}".format(book.sheet_names()))
#sh = book.sheet_by_index(0)
#print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))

#for rx in range(sh.nrows):
#    print(sh.row(rx))
