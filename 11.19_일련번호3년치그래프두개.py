import sqlite3
import xlrd
import xlsxwriter
import math
import scipy as sp
import matplotlib.pyplot as plt
import numpy as np
import datetime

# testDB에 jejutourist2가 이미 있는 경우 아래 코드 주석 해제

# con = sqlite3.connect("testDB")
# cur = con.cursor()
# cur.execute("drop table jejutourist2")
# print("DB table removed !")
# con.commit()

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

#alldata
print("----------- select * from jejutourist2; ------------------------")
for row in cur.execute("select * from jejutourist2;"):
    print(row)
    
#alldata,type
print("----------- select sum(d_individual), sum(d_group), sum(d_leisure) as dum_of_d_group from jejutourist2; ---")
for row in cur.execute("select sum(d_individual), sum(d_group), sum(d_leisure) as dum_of_d_group from jejutourist2;"):
    print("sum(d_individual), sum(d_group), sum(d_leisure)")
    print(row)

#2016,type
print("----------- select month, sum(d_individual), sum(d_group), sum(d_leisure) as dum_of_d_group ---")
print("----------- from jejutourist2_year:2016 group by month order by month ---- ")
for row in cur.execute("select month, sum(d_individual), sum(d_group), sum(d_leisure) as dum_of_d_group \
                        from jejutourist2\
                        where year=2016\
                        group by month\
                        order by month;"):
    print(row)
    
#2016,country
print("----------- select month, sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), ...------")
print("----------- from jejutourist2_year2016 group by month order by month; -----")
for row in cur.execute("select month, sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), \
                        	sum(f_malasia), sum(f_asia_etc), sum(f_usa),\
				sum(f_etc) as dum_of_d_group \
                        from jejutourist2\
                        where year=2016\
                        group by month\
                        order by month;"):
    print(row)
  
#Cumulative
print("---------------------------------------------------------")
print("연도별 누적합을 구합니다.")
con = sqlite3.connect("testDB")
cur = con.cursor()
a=input("년도를 입력하세요 :")
b=input("월을 입력하세요 :")
print("[년도,월,당월방문객,누적방문객]")
sql='select year, month, (d_leisure), (select sum(d_leisure) from jejutourist2 as T where T.year=S.year and T.month<=S.month) as cumul\
                        from jejutourist2 as S\
                        where year=? and month=?\
                        order by year, month'
cur.execute(sql,(a,b))
rows=cur.fetchall()
print(rows)


#해당 국가의 방문객이 많은 월
print("---------------------------------------------------------")
print("해당 국가의 방문객이 많은 월(月)은?")
a=input("국가를 입력하세요(japan,china,hongkong,taiwan,singapore,malasia,asia_etc,usa,etc :")
print("[년도,월,당월방문객]")
if a=='japan':
    for row in cur.execute("select year,month, max(f_japan) as dum_of_d_group \
                            from jejutourist2"):
        print(row)
elif a=='china':
    for row in cur.execute("select year,month, max(f_china) as dum_of_d_group \
                            from jejutourist2"):
        print(row)
elif a=='hongkong':
    for row in cur.execute("select year,month, max(f_hongkong) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)
elif a=='taiwan':
    for row in cur.execute("select year,month, max(f_taiwan) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)
elif a=='singapore':
    for row in cur.execute("select year,month, max(f_singapore) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)
elif a=='malasia':
    for row in cur.execute("select year,month, max(f_malasia) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)
elif a=='asia_etc':
    for row in cur.execute("select year,month, max(f_malasia) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)
elif a=='usa':
    for row in cur.execute("select year,month, max(f_usa) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)
else:
    for row in cur.execute("select year,month, max(f_etc) as dum_of_d_group \
                            from jejutourist2;"):
        print(row)

''' 아직못함
#최대인원이 방문하는 국가?인원수?
'''

#개인 방문객이 가장 많이 방문한 월과 인원수?
print("---------------------------------------------------------")
print("해당 년도의 개인 방문객이 가장 많이 방문한 월(月)과 그때의 방문객 수는?")
a=input("연도를 입력하세요(2016,2017,2018):")
print("[월,당월방문객]")
if a=='2016':
    for row in cur.execute("select month, sum(d_individual) as dum_of_d_group\
                            from jejutourist2\
                            where year=2016"):
        print(row)
elif a=='2017':
    for row in cur.execute("select month, sum(d_individual) as dum_of_d_group \
                            from jejutourist2\
                            where year=2017"):
        print(row)
else:
    for row in cur.execute("select month, sum(d_individual) as dum_of_d_group \
                            from jejutourist2\
                            where year=2018"):
        print(row)


#plot
con = sqlite3.connect("testDB")
cur = con.cursor()
print("=== 월 ====== 단체 ===== 그룹 =====레저====")
cur.execute("select id, d_individual, d_group, d_leisure as dum_of_d_group \
                from jejutourist2\
                order by id;") 
rows = cur.fetchall()
print(rows)

fig=plt.figure()
a=fig.add_subplot(1,2,1)
nda = np.asarray(rows)
x1_1 = nda[:,0]
y1_1 = nda[:,1]
y2_1 = nda[:,2]
y3_1 = nda[:,3]
plt.title('Graph for the number of Individual, Group, Leisure')
plt.xlabel('Month')
a.xaxis.set_ticks(np.arange(1, 37, 1))
a.set_xticklabels(["2016.01","2016.02","2016.03","2016.04","2016.05",\
                      "2016.06","2016.07","2016.08","2016.09","2016.10",\
                      "2016.11","2016.12","2017.01","2017.02","2017.03",\
                      "2017.04","2017.05","2017.06","2017.07","2017.08",\
                      "2017.09","2017.10","2017.11","2017.12","2018.01",\
                      "2018.02","2018.03","2018.04","2018.05","2018.06",\
                      "2018.07","2018.08","2018.09","2018.10","2018.11",\
                      "2018.12"],rotation=30,fontsize="small")
plt.ylabel('No. Visitors')
plt.plot(x1_1, y1_1, '-ro', label="Individual")
plt.plot(x1_1, y2_1, '-.mo', label="Group")
plt.plot(x1_1, y3_1, ':go', label="Leisure")
plt.legend(loc='upper left')


print("=== 월 ======  Chinese =====  ASIA  ========= USA ====")
cur.execute("select id, f_china,f_asia_etc,f_usa \
                from jejutourist2\
                order by id;") 
rows = cur.fetchall()
print(rows)

b=fig.add_subplot(1,2,2)
nda = np.asarray(rows)
x2_2 = nda[:,0]
y1_2 = nda[:,1]
y2_2 = nda[:,2]
y3_2 = nda[:,3]
plt.title('Graph for the number of China, ASIA_etc, USA')
plt.xlabel('Month')
b.xaxis.set_ticks(np.arange(1, 37, 1))
b.set_xticklabels(["2016.01","2016.02","2016.03","2016.04","2016.05",\
                    "2016.06","2016.07","2016.08","2016.09","2016.10",\
                    "2016.11","2016.12","2017.01","2017.02","2017.03",\
                    "2017.04","2017.05","2017.06","2017.07","2017.08",\
                    "2017.09","2017.10","2017.11","2017.12","2018.01",\
                    "2018.02","2018.03","2018.04","2018.05","2018.06",\
                    "2018.07","2018.08","2018.09","2018.10","2018.11",\
                    "2018.12"],rotation=30,fontsize="small")
plt.ylabel('No. Visotors')
plt.plot(x2_2, y1_2, '-bo', label="Chinese")
plt.legend(loc='upper left')
plt.plot(x2_2, y2_2, '-.mo', label="ASIA_ect")
plt.legend(loc='upper left')
plt.plot(x2_2, y3_2, ':go', label="USA")
plt.legend(loc='upper left')
plt.show()

