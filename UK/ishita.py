import xlrd,datetime,xlsxwriter
file_location="D:\Download\Historic COVID-19 Dashboard Data.xlsx"  #EDIT
workbook=xlrd.open_workbook(file_location)
sheet1,sheet2,sheet3,sheet4,sheet5,sheet6=workbook.sheet_by_index(1),workbook.sheet_by_index(2),workbook.sheet_by_index(3),workbook.sheet_by_index(4),workbook.sheet_by_index(5),workbook.sheet_by_index(6)
# print(sheet1.cell_value(9,0))
def sh1(val):         #FOR SHEET1 DATA
    return dates[val],cases[val],cumulative_cases[val]

dates,cases,cumulative_cases=[],[],[]
for i in range(9,sheet1.nrows):
    temp = sheet1.cell_value(i,0)
    temp=xlrd.xldate_as_tuple(temp,0)
    time=str(temp[2])+"/"+str(temp[1])+"/"+str(temp[0])
    dates.append(time)
    cases.append(sheet1.cell_value(i,1))
    cumulative_cases.append(sheet1.cell_value(i,2))
sheet1_data=[[0]*3 for x in range(len(dates)+1)]
sheet1_data[0][0]="Date"
sheet1_data[0][1]="Cases"
sheet1_data[0][2]="Cumulative Cases"
for i in range(1,len(dates)+1):
    sheet1_data[i][0]=dates[i-1]
    sheet1_data[i][1]=cases[i-1]
    sheet1_data[i][2]=cumulative_cases[i-1]

#SHEET 2 UK DEATHS
l=[[0]*7 for x in range(sheet2.nrows-7)]
# print(len(l))
l[0][0],l[0][1],l[0][2],l[0][3],l[0][4],l[0][5],l[0][6]="Date","Deaths","UK","ENGLAND","Scotland","Wales","NORTHERN IRELAND"
for i in range(1,sheet2.nrows-7):
    temp = sheet2.cell_value(i+7,0)
    temp=xlrd.xldate_as_tuple(temp,0)
    time=str(temp[2])+"/"+str(temp[1])+"/"+str(temp[0])
    l[i][0]=time
    l[i][1]=sheet2.cell_value(i+7,1)
    l[i][2]=sheet2.cell_value(i+7,2)
    l[i][3]=sheet2.cell_value(i+7,3)
    l[i][4]=sheet2.cell_value(i+7,4)
    l[i][5]=sheet2.cell_value(i+7,5)
    l[i][6]=sheet2.cell_value(i+7,6)

sheet3_data=[[0]*sheet3.ncols for x in range(6)]
for i in range(1,6):
    for j in range(0,sheet3.ncols):
        sheet3_data[i][j]=sheet3.cell_value(i+7,j)
sheet3_data[0][0]="Area Code"
sheet3_data[0][1]="Area Name"
for i in range(2,sheet3.ncols):
    temp = sheet3.cell_value(7,i)
    temp=xlrd.xldate_as_tuple(temp,0)
    time=str(temp[2])+"/"+str(temp[1])+"/"+str(temp[0])
    sheet3_data[0][i]=time
# print(sheet3_data)
sheet3_date=sheet3_data[0][2:]
# print(sheet3_date)
sheet3_eng_cases=sheet3_data[1][2:]
sheet3_scot_cases=sheet3_data[2][2:]
sheet3_wales_cases=sheet3_data[3][2:]
sheet3_uk_cases=sheet3_data[4][2:]
sheet3_data_req=[]
sheet3_data_req.append(['Area','Date','Cases'])
for i in range(len(sheet3_date)):
    sheet3_data_req.append(['England',sheet3_date[i],sheet3_eng_cases[i]])
for i in range(len(sheet3_date)):
    sheet3_data_req.append(['Scotland',sheet3_date[i],sheet3_scot_cases[i]])
for i in range(len(sheet3_date)):
    sheet3_data_req.append(['Wales',sheet3_date[i],sheet3_wales_cases[i]])
for i in range(len(sheet3_date)):
    sheet3_data_req.append(['UK',sheet3_date[i],sheet3_uk_cases[i]])
# print(sheet3_data_req[85])

#SHEET 3 DONE STARTING 4 not complete
sheet4_data=[[0]*(sheet4.ncols-2) for x in range(9)]
sheet4_data[0][0]="Area Name"
sheet4_data[1][0]="Unconfirmed"
sheet4_data[2][0]="London"
sheet4_data[3][0]="South East"
sheet4_data[4][0]="East Of England"
sheet4_data[5][0]="Midlands"
sheet4_data[6][0]="North East and Yorkshire"
sheet4_data[7][0]="North West"
sheet4_data[8][0]="England"
for i in range(0,sheet4.nrows-2):
    temp = sheet4.cell_value(7,i+2)
    temp=xlrd.xldate_as_tuple(temp,0)
    time=str(temp[2])+"/"+str(temp[1])+"/"+str(temp[0])
    sheet4_data[0][i+1]=time
for i in range(1,9):
    for j in range(1,sheet4.ncols-2):
        sheet4_data[i][j]=sheet4.cell_value(i+7,j+2)
# for i in sheet4_data:
#     print(i)
#SHEET 5 start->
sheet5_data=[[0]*sheet5.ncols for x in range(sheet5.nrows-7)]
for i in range(1,len(sheet5_data)):
    for j in range(0,sheet5.ncols):
        sheet5_data[i][j]=sheet5.cell_value(i+7,j)
sheet5_data[0][0]="AREA CoDE"
sheet5_data[0][1]="Area Name"
for i in range(2,sheet5.ncols):
    temp = sheet4.cell_value(7,i)
    temp=xlrd.xldate_as_tuple(temp,0)
    time=str(temp[2])+"/"+str(temp[1])+"/"+str(temp[0])
    sheet5_data[0][i]=time
# print(sheet5_data[0])
#sheet 5 done
#sheet 6->
sheet6_data=[[0,0] for n in range(sheet6.nrows-6)]
sheet6_data[0][0]='Date'
sheet6_data[0][1]='Cumulative Counts'
for i in range(1,len(sheet6_data)):
    temp = sheet6.cell_value(i+6,0)
    temp=xlrd.xldate_as_tuple(temp,0)
    time=str(temp[2])+"/"+str(temp[1])+"/"+str(temp[0])
    sheet6_data[i][0]=time
    sheet6_data[i][1]=sheet6.cell_value(i+6,1)
# print(sheet6_data)

workbook=xlsxwriter.Workbook(r'D:\Download\Data.xlsx') #EDIT 
worksheet=workbook.add_worksheet()

worksheet1=workbook.add_worksheet()
for i in range(0,len(sheet1_data)):
    for j in range(0,3):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet1.write(cor,str(sheet1_data[i][j]))
worksheet2=workbook.add_worksheet()
for i in range(0,len(l)):
    for j in range(0,len(l[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet2.write(cor,str(l[i][j]))
worksheet3=workbook.add_worksheet()
for i in range(0,len(sheet3_data)):
    for j in range(0,len(sheet3_data[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet3.write(cor,str(sheet3_data[i][j]))
worksheet4=workbook.add_worksheet()
for i in range(0,len(sheet4_data)):
    for j in range(0,len(sheet4_data[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet4.write(cor,str(sheet4_data[i][j]))
worksheet5=workbook.add_worksheet()
for i in range(0,len(sheet5_data)):
    for j in range(0,len(sheet5_data[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet5.write(cor,str(sheet5_data[i][j]))

worksheet.write('A1','Date')
worksheet.write('B1','Cumulative Counts')
for i in range(1,len(sheet6_data)):
    a1='A'+str(i+1)
    worksheet.write(a1,str(sheet6_data[i][0]))
    b1='B'+str(i+1)
    worksheet.write(b1,str(sheet6_data[i][1]))

workbook.close()