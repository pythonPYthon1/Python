# danh sanh person type 1
list_type1 = ['Person_1', 'Person_2', 'Person_3', 'Person_4', 'Person_5', 'Person_6']
len1 = int(len(list_type1))
# danh sach person type 2
list_type2 = ['Person_A', 'Person_B', 'Person_C', 'Person_D']
len2 = int(len(list_type2))
# danh sach person type 3
list_type3 = ['DuongPT3']
len3 = int(len(list_type3))

import numpy 
import xlwt
import openpyxl
from xlwt import Workbook
import mysql.connector

# ham thay doi vi tri cac value di 1 don vi 
def roll1(lst):
    return numpy.roll(lst, 1)

# ham sap xep lich
def schedule(day, list_type1, list_type2, list_type3):
    # tao file excel
    wb = openpyxl.Workbook()
    sheet = wb.active                       
    sheet.merge_cells('A1:J1')
    sheet['A1'] = 'LICH TRUC GIAM SAT ATTT'
    sheet.merge_cells('A2:B2')
    sheet['A2'] = 'THOI GIAN: 30 NGAY'
    sheet.merge_cells('A3:A5')
    sheet['A3'] = 'Ngay'
    sheet.merge_cells('B3:J3')
    sheet['B3'] = 'Ca truc'
    sheet.merge_cells('B4:D4')
    sheet['B4'] = 'Ca 1 (0:00 - 8:00)'
    sheet.merge_cells('E4:G4')
    sheet['E4'] = 'Ca 2 (8:00 - 16:00)'
    sheet.merge_cells('H4:J4')
    sheet['H4'] = 'Ca 3 (16:00 - 0:00)'
    sheet['B5'] = 'Tier 1'
    sheet['C5'] = 'Tier 2'
    sheet['D5'] = 'Tier 3'
    sheet['E5'] = 'Tier 1'
    sheet['F5'] = 'Tier 2'
    sheet['G5'] = 'Tier 3'
    sheet['H5'] = 'Tier 1'
    sheet['I5'] = 'Tier 2'
    sheet['J5'] = 'Tier 3'

    # connect to mysql
    mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    passwd="longlnhe")
    mycursor = mydb.cursor()
    mycursor.execute("USE TestDB")
    #mycursor.execute("CREATE DATABASE Schedule")
    mycursor.execute("DROP TABLE IF EXISTS TASK_SCHEDULE")
    mycursor.execute("CREATE TABLE TASK_SCHEDULE (id INT(11) NOT NULL AUTO_INCREMENT, name VARCHAR(45) NOT NULL, timestamp DATETIME , shift VARCHAR(45) NOT NULL, PRIMARY KEY(id)) ENGINE=InnoDB AUTO_INCREMENT=397 DEFAULT CHARSET=latin1")
    
    count = 0
    a = 0
    b = 0
    i = 0
    line = ''
    while i < int(day):
        # 1 ngay co 3 ca:
        for j in range (3):
            # 1 ca
            line1 =  list_type1[a] + " " + list_type2[b] + " "  + list_type3[0] + " "
            mycursor.execute("INSERT INTO TASK_SCHEDULE(name, shift) VALUES('"+ list_type1[a] +"', '"+ str((j + 1))+"')")
            
            # 1 ngay = 3 ca cong lai
            line += line1
            count += 1
            # khi count == 6 can phai roll lai type 1 de tranh truong hop 1 nguoi type 1 lam slot 1 2 lan lien tiep
            if count == 6:
                list_type1 = roll1(list_type1)
                # reset lai a va bien count
                a = 0
                count = 0
            else:
                # neu count != 6 thi tang a len 1
                a = a + 1
        # khi type 2 dc goi het thì se goi lai tu dau
        if (b == 3):
            b = 0
        else:
            b = b + 1
        mycursor.execute("INSERT INTO TASK_SCHEDULE(name, shift) VALUES('"+ list_type2[b] +"', '1 2 3')")
        mycursor.execute("INSERT INTO TASK_SCHEDULE(name, shift) VALUES('"+ list_type3[0] +"', '1 2 3')")
        #print("Day " + str(i + 1) + ": " + line)
        # ghi vao excel
        # sheet.cell(row = i + 5, column = 1).value = 
        array = line.split(' ')
        for k in range (len(array)):
            sheet.cell(row = i + 6, column = k + 2).value =  array[k]
            
        i = i + 1
        line = ''
    mydb.commit()
    wb.save('Ex1.xlsx')
  

schedule(30, list_type1, list_type2, list_type3)

