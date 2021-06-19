import sqlite3
import xlrd
import xlsxwriter
import math
import scipy as sp
import matplotlib.pyplot as plt
import numpy as np
import datetime

# testDB에 jejutourist2가 이미 있는 경우 아래 코드 주석 해제

con = sqlite3.connect("testDB")
cur = con.cursor()
cur.execute("drop table jejutourist2")
print("DB table removed !")
con.commit()

con = sqlite3.connect("testDB")
cur = con.cursor()
cur.execute("create table jejutourist2(\
    id int,\
    year int,\
    month int,\
    d_individual int, \
    d_group int,\
    d_leisure int,\
    d_conference int,\
    d_rest int,\
    d_famliy int,\
    d_education int,\
    d_etc int,\
    f_japan int,\
    f_china int,\
    f_hongkong int,\
    f_taiwan int,\
    f_singapore int,\
    f_malasia int,\
    f_asia_etc int,\
    f_usa int,\
    f_etc int)"
)
print("DB table created !")
con.commit()

con = sqlite3.connect("testDB")
cur = con.cursor()

wb = xlrd.open_workbook('C:/data/jejuall.xls')
print("The number of worksheets is", wb.nsheets)
print("Worksheet name(s):", wb.sheet_names())
cur.execute("delete from jejutourist2")

i=0
while i < wb.nsheets:
    sh = wb.sheet_by_index(i)
    print(sh.name, sh.nrows, sh.ncols)
    cur.execute("insert into jejutourist2 values (?,?,?,?,?,?,\
                                           ?,?,?,?,?,\
                                           ?,?,?,?,?,\
                                           ?,?,?,?)", 
                (i+1,\
                 (((str(sh.cell(2,6))).split("'"))[1])[:4],\
                 (((str(sh.cell(0,0))).split())[1]).rstrip("월"),\
                 ((str(sh.cell(7,6))).split(":"))[1],\
                 ((str(sh.cell(9,6))).split(":"))[1],\
                 ((str(sh.cell(22,6))).split(":"))[1],\
                 ((str(sh.cell(24,6))).split(":"))[1],\
                 ((str(sh.cell(26,6))).split(":"))[1],\
                 ((str(sh.cell(28,6))).split(":"))[1],\
                 ((str(sh.cell(30,6))).split(":"))[1],\
                 ((str(sh.cell(32,6))).split(":"))[1],\
                 ((str(sh.cell(36,6))).split(":"))[1],\
                 ((str(sh.cell(38,6))).split(":"))[1],\
                 ((str(sh.cell(40,6))).split(":"))[1],\
                 ((str(sh.cell(42,6))).split(":"))[1],\
                 ((str(sh.cell(44,6))).split(":"))[1],\
                 ((str(sh.cell(46,6))).split(":"))[1],\
                 ((str(sh.cell(54,6))).split(":"))[1],\
                 ((str(sh.cell(56,6))).split(":"))[1],\
                 ((str(sh.cell(58,6))).split(":"))[1]
                ))
    i = i+1    
con.commit()

print("---------------------------------------------------------")
print("====년도별 최대 인원이 방문하는 국가====")
a=input("년도를 입력하세요(2016, 2017, 2018) :")
if a=='2016':
    for row in cur.execute("select year,max(sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa)) as dum_of_d_group \
                            from jejutourist2\
                            where year=2016;"):
        print(row)
      
elif a=='2017':
    for row in cur.execute("select year, max(sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa)) as dum_of_d_group \
                            from jejutourist2\
                            where year=2017;"):
        print(row)
        
else:
    for row in cur.execute("select year, max(sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa)) as dum_of_d_group \
                            from jejutourist2\
                            where year=2018;"):
        print(row)

#plot

con = sqlite3.connect("testDB")
cur = con.cursor()
print("=== 월 ====== 단체 ===== 그룹 =====레저====")
fig=plt.figure()
a=fig.add_subplot(1,3,1)
cur.execute("select sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa) as dum_of_d_group \
                            from jejutourist2\
                            where year=2016;")
rows = cur.fetchall()
print(rows)
ratio = rows
labels = ['japan','china','hongkong','taiwan','singapore','malasia','usa']
plt.pie(ratio, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)

b=fig.add_subplot(1,3,2)
cur.execute("select sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa) as dum_of_d_group \
                            from jejutourist2\
                            where year=2017;")
rows = cur.fetchall()
print(rows)
ratio = rows
labels = ['japan','china','hongkong','taiwan','singapore','malasia','usa']
plt.pie(ratio, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)

c=fig.add_subplot(1,3,3)
cur.execute("select sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa) as dum_of_d_group \
                            from jejutourist2\
                            where year=2018;")
rows = cur.fetchall()
print(rows)
ratio = rows
labels = ['japan','china','hongkong','taiwan','singapore','malasia','usa']
plt.pie(ratio, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)
plt.show()

