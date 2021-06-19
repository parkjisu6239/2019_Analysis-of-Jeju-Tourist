



# 제주 입도 현황 DB 화 및 시각화

>  2019.11.19 충북대학교 수학과 박지수
>
> 수업명 : 고급 프로그래밍 언어
>
> 지도교수 : 조완섭



## 목적

- 공시용 엑셀 데이터에서 데이터를 추출하여 분석 및 시각화 하기 좋은 형태로 구성
- 제주 관광 특징을 조사하고, 향후 계획 수립에 기여



## 개요

- Excel에서 필요한 데이터를 추출하여  DB로 저장한다. 
- 저장한 데이터를 바탕으로 사용자 입력에 따라 plot을 출력한다
  - 파이차트, 히스토그램, 선형 plot등 다양한 형태로 제공





## 상세 내용



#### 1. excel 원본 데이터 확인

- 월별로 시트로 다르게 구성되어 있다.
- 전체적인 흐름을 보기 어려운 구조
- 연별, 월별 데이터를 파악하기 힘들다.

![image-20210619150955764](.\README.assets\image-20210619150955764.png)





#### 2. DB 저장

```python
con = sqlite3.connect("testDB") # DB 생성
cur = con.cursor() # DB 연결
cur.execute("create table jejutourist2(\ # 테이블, 스키마 생성
    ...
)

wb = xlrd.open_workbook('C:/data/jejuall.xls') # 엑셀 열기

# 엑셀 -> DB insert
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
                 ...
```



#### 3. 옵션 선택

- 연도 선택
- 연간 국가별 관광객  중 최대값 추출

```python
if a=='2016':
    for row in cur.execute("select year,max(sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa)) as dum_of_d_group \
                            from jejutourist2\
                            where year=2016;"):
```



#### 4. 시각화

```python
fig=plt.figure() # plot 초기 생성
a=fig.add_subplot(1,3,1) # subplot
cur.execute("select sum(f_japan), sum(f_china), sum(f_hongkong), sum(f_taiwan), sum(f_singapore), sum(f_malasia), sum(f_usa) as dum_of_d_group \
                            from jejutourist2\
                            where year=2016;")
rows = cur.fetchall() # 위에서 추가한 모든 데이터 가져오기
print(rows)
ratio = rows
labels = ['japan','china','hongkong','taiwan','singapore','malasia','usa']
plt.pie(ratio, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)
```



#### 5. 결과

- 2016, 2017, 2018 연도별 국가별 파이차트

![image-20210619152252457](.\README.assets\image-20210619152252457.png)



- 2016~2018 타입(개인, 그룹, 레저) 별 관광객 현황

![image-20210619152425518](.\README.assets\image-20210619152425518.png)



- (좌) 월별, 타입별 관광객 누적 (우) 월별, 국가별 관광객 누적

![image-20210619152519305](.\README.assets\image-20210619152519305.png)



- 월별 주요 국가 관광객 수

![image-20210619152621580](.\README.assets\image-20210619152621580.png)



#### 6. 분석

- 중국 관광객이 2017.04 를 기준으로 급감했다. 이는 당시 사드배치의 영향으로 중국인 관광객이 감소한 것으로 해석할 수 있다.

  ![image-20210619153823042](.\README.assets\image-20210619153823042.png)

- 개인 관광객이 꾸준히 증가하고 있다. 1인 혹은 2인 등 소규모 관광객이 즐길 수 있는 여행 코스나 식당을 적극적으로 홍보하는 것이 좋을 것으로 예상 된다.

- 오히려 한여름 성수기보다 봄 가을, 심지어 겨울에 관광객이 많다. 계절별 관광 코스를 구성하여 관람객의 니즈를 충족 시키는 것이 좋다.





## 결론 및 배운점

- 데이터의 형태에 따라 데이터의 가치가 달라지기도 한다.
- 데이터를 보다 효율적으로 사용하기 위해서 DB 형태로 저장하고, 이를 분석하면 향후 관광 산업에 기여할 수 있을 것이다.
- **xlrd** 로 엑셀 파일을 열고, 셀에 접근할 수 있다.
- **sqlite3** 로 sql문을 쉽게 사용할 수 있다.