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
import numpy as np
import pandas as pd
import xlrd
import xlwt
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment,PatternFill

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
for item in all_input_files:
    print(item)

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

                dst_dir="output\\output_bo\\bo_kqht_"+ lop + "_" + str(now) + ".xlsx"

                #Xoa truoc 2 dong dau cua dataframe
                df = df.iloc[2:]

                writer = pd.ExcelWriter(dst_dir,engine = "openpyxl", mode = 'w')
                df.to_excel(writer,sheet_name = 'Sheet1', startrow=2, startcol=0, header=False, index=False)

                
                writer.save()
                
                #Xu ly tieu de bang openpyxl
                wb = openpyxl.load_workbook(dst_dir)
                ws = wb.active

                font_title = openpyxl.styles.Font(name='Times New Roman',size = 11,bold=True,color= '00FF0000')
                font_marktitle = openpyxl.styles.Font(name='Times New Roman',size = 11,bold=True)
                style_title = openpyxl.styles.Alignment(horizontal='center', vertical='center',wrap_text= True)
                
                grey = "DDDDDD"
                black= "00808080"
                

                

                def headerDealer(col,row,strmerge,strip,fontchoice,backgroundchoice):
                    cell_dealing = ws.cell(column=col,row=row)
                    cell_dealing.value = strip
                    cell_dealing.font = fontchoice
                    ws.merge_cells(strmerge)
                    cell_dealing.alignment = style_title
                    cell_dealing.fill =  PatternFill("solid", fgColor=backgroundchoice)

                headerDealer(1,1,"A1:A2","STT",font_title,grey)
                headerDealer(2,1,"B1:B2","Mã lớp",font_title,grey)
                headerDealer(3,1,"C1:C2","Mã học sinh",font_title,grey)
                headerDealer(4,1,"D1:D2","Họ tên",font_title,grey)
                headerDealer(5,1,"E1:E2","Ngày sinh",font_title,grey)
                headerDealer(6,1,"F1:F2","Toán",font_marktitle,black)
                headerDealer(7,1,"G1:G2","Vật lý",font_marktitle,black)
                headerDealer(8,1,"H1:H2","Hóa học",font_marktitle,black)
                headerDealer(9,1,"I1:I2","Sinh học",font_marktitle,black)
                headerDealer(10,1,"J1:J2","Tin học",font_marktitle,black)
                headerDealer(11,1,"K1:K2","Ngữ văn",font_marktitle,black)
                headerDealer(12,1,"L1:L2","Lịch sử",font_marktitle,black)
                headerDealer(13,1,"M1:M2","Địa lý",font_marktitle,black)
                headerDealer(14,1,"N1:N2","Ngoại ngữ",font_marktitle,black)
                headerDealer(15,1,"O1:O2","Công nghệ",font_marktitle,black)
                headerDealer(16,1,"P1:P2","GD QP-AN",font_marktitle,black)
                headerDealer(17,1,"Q1:Q2","Thể dục",font_marktitle,black)
                headerDealer(18,1,"R1:S1","Tự chọn",font_marktitle,black)
                headerDealer(20,1,"T1:T2","GDCD",font_marktitle,black)
                headerDealer(21,1,"U1:U2","ĐTB các môn",font_marktitle,black)
                headerDealer(22,1,"V1:X1","Kết quả xếp loại và DH thi đua",font_marktitle,black)


                def withoutmerge_headerDealer(col,row,strip,fontchoice,backgroundchoice):
                    cell_dealing = ws.cell(column=col,row=row)
                    cell_dealing.value = strip
                    cell_dealing.font = fontchoice
                    cell_dealing.alignment = style_title
                    cell_dealing.fill =  PatternFill("solid", fgColor=backgroundchoice)

                withoutmerge_headerDealer(18,2,"Ngoại ngữ 2",font_marktitle,black)
                withoutmerge_headerDealer(19,2,"Nghề phổ thông",font_marktitle,black)
                withoutmerge_headerDealer(22,2,"Học lực",font_marktitle,black)
                withoutmerge_headerDealer(23,2,"Hạnh kiểm",font_marktitle,black)
                withoutmerge_headerDealer(24,2,"Danh hiệu thi đua",font_marktitle,black)


                


                #Def dich chuyen du lieu
                def move_range(startcol,startrow,endcol,endrow,rowmove,colmove):
                    str_range = startcol+str(startrow)+':'+endcol+str(endrow)
                    ws.move_range(str_range,rows=rowmove,cols=colmove)

                #Dich chuyen phan du lieu dataframe
                move_range('C',3,'S',final_row+1,0,3)
                move_range('B',3,'B',final_row+1,0,2)
                move_range('T',3,'V',final_row+1,0,2)
                move_range('S',3,'S',final_row+1,0,2)
                move_range('O',3,'O',final_row+1,0,5)


                #Dat lai cot dtb lay 1 so sau dau phay
                for row in range(1, final_row+2):
                    ws["F{}".format(row)].number_format = '#,#0.0'
                    ws["G{}".format(row)].number_format = '#,#0.0'
                    ws["H{}".format(row)].number_format = '#,#0.0'
                    ws["I{}".format(row)].number_format = '#,#0.0'
                    ws["J{}".format(row)].number_format = '#,#0.0'
                    ws["K{}".format(row)].number_format = '#,#0.0'
                    ws["L{}".format(row)].number_format = '#,#0.0'
                    ws["M{}".format(row)].number_format = '#,#0.0'
                    ws["N{}".format(row)].number_format = '#,#0.0'
                    ws["O{}".format(row)].number_format = '#,#0.0'
                    ws["P{}".format(row)].number_format = '#,#0.0'
                    ws["R{}".format(row)].number_format = '#,#0.0'
                    ws["S{}".format(row)].number_format = '#,#0.0'
                    ws["T{}".format(row)].number_format = '#,#0.0'
                    ws["U{}".format(row)].number_format = '#,#0.0'
                    


                #st_range = 'C3:S'+ str(final_row+1)
                #print(st_range)
                #ws.move_range(st_range,rows=0,cols=3)

                #name_range = 'B3:B'+str(final_row+1)
                #ws.move_range(name_range,rows=0,cols=2)

                #kqxl_range = 'T3:V'+str(final_row+1)
                #ws.move_range(kqxl_range,rows=0,cols=2)

                #kqxl_range = 'S3:S'+str(final_row+1)
                #ws.move_range(kqxl_range,rows=0,cols=2)






                #Dong khung 
                def __format_ws__(ws, cell_range):
                    border = Border(left=Side(border_style='thin', color='000000'),
                        right=Side(border_style='thin', color='000000'),
                        top=Side(border_style='thin', color='000000'),
                        bottom=Side(border_style='thin', color='000000'))

                    rows = ws[cell_range]
                    for row in rows:
                        for cell in row:
                            cell.border = border

                cells_range = 'A1:X'+str(final_row+1)
                __format_ws__(ws,cells_range)




                wb.save(dst_dir)
                






                #atest = ws['A8']
                #print(atest.value)

                

#Mau xu ly
#book = xlrd.open_workbook("input\Sodiem_Tonghop_10A1.xls")
#print("The number of worksheets is {0}".format(book.nsheets))
#print("Worksheet name(s): {0}".format(book.sheet_names()))
#sh = book.sheet_by_index(0)
#print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))

#for rx in range(sh.nrows):
#    print(sh.row(rx))


#Tu lieu xu ly cho file cua so 

  #src_dir0="Not_touch_core\\so_nhapdiemchitiet.xls"
                #src_dir="Not_touch_core\\so_nhapdiemchitiet.xlsx"
                #dst_dir="output\\output_so\\so_nhapdiemchitiet_"+ lop + "_" + str(now) + ".xlsx"

#str_school_title = 'NHẬP ĐIỂM CHI TIẾT MÔN LỚP ' + lop
                
                #style_header = openpyxl.styles.Font(name='Times New Roman',sz=14,b=True)
                #style_center = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                #normalstyle = openpyxl.styles.Font(name='Times New Roman')

                #wb = openpyxl.load_workbook(dst_dir)
                #ws = wb.active

                #cell_so = ws.cell(row=1,column=1)
                #cell_so.value = 'Sở giáo dục và đào tạo'
                #cell_so.font = normalstyle
                #cell_truong = ws.cell(row=2,column=1)
                #cell_truong.value = 'Đơn vị: THPT Đức Hòa'
                #cell_truong.font = normalstyle
                

                ##Title lop
                #cell_title = ws.cell(row=4,column=1)
                #cell_title.value = str_school_title
                #cell_title.font = style_header
                #cell_title.alignment = style_center
                #ws.merge_cells("A4:M4")
                #wb.save(dst_dir)













#Raw procedures

#Xu ly tieu de
                #df_upper_title = pd.DataFrame({'Data':['Sở giáo dục và đào tạo','Đơn vị: THPT Đức Hòa']})
                #df_upper_title.to_excel(writer, startcol=0,startrow=0, header=None, index=False)
                #str_school_title = 'NHẬP ĐIỂM CHI TIẾT MÔN LỚP ' + lop

                #str_school_title = 'NHẬP ĐIỂM CHI TIẾT MÔN LỚP ' + lop
                #list_sc_title = list([str_school_title]*19)
                #print(list_sc_title)
                #df_class = pd.DataFrame([list_sc_title])
             
                #Dang bi vuong tieu de lop
                #Style option: 'color': font-weight: bold lua chon de xu ly bold face
                #df_class.style.set_properties(color="yellow")
                
                #Sap vi tri tieu de 
                #df_class.style.set_properties(**{'font-size':'14pt', 'text-align':'center','font-weight':'bold'}).to_excel(writer, startcol=0, startrow=3, header = None, index=False)
                
                #Luu du lieu lai