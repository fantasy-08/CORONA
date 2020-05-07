# -*- coding: utf-8 -*-
"""
Created on Wed Apr  1 15:19:25 2020

@author: ishita
"""
import os
import boto3
import xlrd,xlsxwriter
#import urllib
import requests
#import pandas as pd
#os.getcwd()


dls = "https://fingertips.phe.org.uk/documents/Historic%20COVID-19%20Dashboard%20Data.xlsx"
resp = requests.get(dls)
current_directory = os.getcwd()
UK_base_data_path = os.path.join(current_directory, r'COVID_UK_DATA_PUBLIC_HEALTH.xls')
if os.path.exists(UK_base_data_path):
    os.remove(UK_base_data_path) 
else:
    pass  
with open('COVID_UK_DATA_PUBLIC_HEALTH.xls', 'wb') as output:
    output.write(resp.content)

file_location= r"COVID_UK_DATA_PUBLIC_HEALTH.xls"
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

#sheet 2 stored in l[][] variable
    
#Sheet 3
    
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
sheet3_ni_cases=sheet3_data[5][2:]
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
for i in range(len(sheet3_date)):
    sheet3_data_req.append(['Northern Ireland',sheet3_date[i],sheet3_ni_cases[i]])
# print(sheet3_data_req[85])
        
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
    
sheet5_date=sheet5_data[0][2:]

sheet5_Hartlepool_cases=sheet5_data[1][2:]
sheet5_Middlesbrough_cases=sheet5_data[2][2:]
sheet5_RC_cases=sheet5_data[3][2:]
sheet5_Tees_cases=sheet5_data[4][2:]
sheet5_Darlington_cases=sheet5_data[5][2:]
sheet5_Halton_cases=sheet5_data[6][2:]
sheet5_Warrington_cases=sheet5_data[7][2:]
sheet5_Blackburn_cases=sheet5_data[8][2:]
sheet5_Blackpool_cases=sheet5_data[9][2:]
sheet5_Kingston_Hull_cases=sheet5_data[10][2:]
sheet5_ERY_cases=sheet5_data[11][2:]
sheet5_NEL_cases=sheet5_data[12][2:]
sheet5_NL_cases=sheet5_data[13][2:]
sheet5_York_cases=sheet5_data[14][2:]
sheet5_Derby_cases=sheet5_data[15][2:]
sheet5_Leicester_cases=sheet5_data[16][2:]
sheet5_Rutland_cases=sheet5_data[17][2:]
sheet5_Nottingham_cases=sheet5_data[18][2:]
sheet5_Herefordshire_cases=sheet5_data[19][2:]
sheet5_Telford_cases=sheet5_data[20][2:]
sheet5_Stoke_cases=sheet5_data[21][2:]
sheet5_Bath_cases=sheet5_data[22][2:]
sheet5_Bristol_cases=sheet5_data[23][2:]
sheet5_Nsomerset_cases=sheet5_data[24][2:]
sheet5_SGloucestershire_cases=sheet5_data[25][2:]
sheet5_Plymouth_cases=sheet5_data[26][2:]
sheet5_Torbay_cases=sheet5_data[27][2:]
sheet5_Swindon_cases=sheet5_data[28][2:]
sheet5_Peterborough_cases=sheet5_data[29][2:]
sheet5_Luton_cases=sheet5_data[30][2:]
sheet5_Southend_cases=sheet5_data[31][2:]
sheet5_Thurrock_cases=sheet5_data[32][2:]
sheet5_Medway_cases=sheet5_data[33][2:]
sheet5_Bracknell_cases=sheet5_data[34][2:]
sheet5_WBerkshire_cases=sheet5_data[35][2:]
sheet5_Reading_cases=sheet5_data[36][2:]
sheet5_Slough_cases=sheet5_data[37][2:]
sheet5_Windsor_cases=sheet5_data[38][2:]
sheet5_Wokingham_cases=sheet5_data[39][2:]
sheet5_Milton_cases=sheet5_data[40][2:]
sheet5_Brighton_cases=sheet5_data[41][2:]
sheet5_Portsmouth_cases=sheet5_data[42][2:]
sheet5_Southampton_cases=sheet5_data[43][2:]
sheet5_Isle_Wight_cases=sheet5_data[44][2:]
sheet5_Durham_cases=sheet5_data[45][2:]
sheet5_CheshireE_cases=sheet5_data[46][2:]
sheet5_CheshireW_cases=sheet5_data[47][2:]
sheet5_Shropshire_cases=sheet5_data[48][2:]
sheet5_Cornwall_cases=sheet5_data[49][2:]
sheet5_Wiltshire_cases=sheet5_data[50][2:]
sheet5_Bedford_cases=sheet5_data[51][2:]
sheet5_CBedfordshire_cases=sheet5_data[52][2:]
sheet5_Northumberland_cases=sheet5_data[53][2:]
sheet5_Bournemouth_cases=sheet5_data[54][2:]
sheet5_Dorset_cases=sheet5_data[55][2:]
sheet5_Bolton_cases=sheet5_data[56][2:]
sheet5_Bury_cases=sheet5_data[57][2:]
sheet5_Manchester_cases=sheet5_data[58][2:]
sheet5_Oldham_cases=sheet5_data[59][2:]
sheet5_Rochdale_cases=sheet5_data[60][2:]
sheet5_Salford_cases=sheet5_data[61][2:]
sheet5_Stockport_cases=sheet5_data[62][2:]
sheet5_Tameside_cases=sheet5_data[63][2:]
sheet5_Trafford_cases=sheet5_data[64][2:]
sheet5_Wigan_cases=sheet5_data[65][2:]
sheet5_Knowsley_cases=sheet5_data[66][2:]
sheet5_Liverpool_cases=sheet5_data[67][2:]
sheet5_Helens_cases=sheet5_data[68][2:]
sheet5_Sefton_cases=sheet5_data[69][2:]
sheet5_Wirral_cases=sheet5_data[70][2:]
sheet5_Barnsley_cases=sheet5_data[71][2:]
sheet5_Doncaster_cases=sheet5_data[72][2:]
sheet5_Rotherham_cases=sheet5_data[73][2:]
sheet5_Sheffield_cases=sheet5_data[74][2:]
sheet5_Newcastle_cases=sheet5_data[75][2:]
sheet5_Ntyneside_cases=sheet5_data[76][2:]
sheet5_STyneside_cases=sheet5_data[77][2:]
sheet5_Sunderland_cases=sheet5_data[78][2:]
sheet5_Birmingham_cases=sheet5_data[79][2:]
sheet5_Coventry_cases=sheet5_data[80][2:]
sheet5_Dudley_cases=sheet5_data[81][2:]
sheet5_Sandwell_cases=sheet5_data[82][2:]
sheet5_Solihull_cases=sheet5_data[83][2:]
sheet5_Walsall_cases=sheet5_data[84][2:]
sheet5_Wolverhampton_cases=sheet5_data[85][2:]
sheet5_Bradford_cases=sheet5_data[86][2:]
sheet5_Calderdale_cases=sheet5_data[87][2:]
sheet5_Kirklees_cases=sheet5_data[88][2:]
sheet5_Leeds_cases=sheet5_data[89][2:]
sheet5_Wakefield_cases=sheet5_data[90][2:]
sheet5_Gateshead_cases=sheet5_data[91][2:]
sheet5_Barking_cases=sheet5_data[92][2:]
sheet5_Barnet_cases=sheet5_data[93][2:]
sheet5_Bexley_cases=sheet5_data[94][2:]
sheet5_Brent_cases=sheet5_data[95][2:]
sheet5_Bromley_cases=sheet5_data[96][2:]
sheet5_Camden_cases=sheet5_data[97][2:]
sheet5_Croydon_cases=sheet5_data[98][2:]
sheet5_Ealing_cases=sheet5_data[99][2:]
sheet5_Enfield_cases=sheet5_data[100][2:]
sheet5_Greenwich_cases=sheet5_data[101][2:]
sheet5_Hackney_cases=sheet5_data[102][2:]
sheet5_Hammersmith_cases=sheet5_data[103][2:]
sheet5_Haringey_cases=sheet5_data[104][2:]
sheet5_Harrow_cases=sheet5_data[105][2:]
sheet5_Havering_cases=sheet5_data[106][2:]
sheet5_Hillingdon_cases=sheet5_data[107][2:]
sheet5_Hounslow_cases=sheet5_data[108][2:]
sheet5_Islington_cases=sheet5_data[109][2:]
sheet5_Kensington_cases=sheet5_data[110][2:]
sheet5_Kingston_Thames_cases=sheet5_data[111][2:]
sheet5_Lambeth_cases=sheet5_data[112][2:]
sheet5_Lewisham_cases=sheet5_data[113][2:]
sheet5_Merton_cases=sheet5_data[114][2:]
sheet5_Newham_cases=sheet5_data[115][2:]
sheet5_Redbridge_cases=sheet5_data[116][2:]
sheet5_Richmond_cases=sheet5_data[117][2:]
sheet5_Southwark_cases=sheet5_data[118][2:]
sheet5_Sutton_cases=sheet5_data[119][2:]
sheet5_Tower_cases=sheet5_data[120][2:]
sheet5_Waltham_cases=sheet5_data[121][2:]
sheet5_Wandsworth_cases=sheet5_data[122][2:]
sheet5_Westminster_cases=sheet5_data[123][2:]
sheet5_Buckinghamshire_cases=sheet5_data[124][2:]
sheet5_Cambridgeshire_cases=sheet5_data[125][2:]
sheet5_Cumbria_cases=sheet5_data[126][2:]
sheet5_Derbyshire_cases=sheet5_data[127][2:]
sheet5_Devon_cases=sheet5_data[128][2:]
sheet5_ESussex_cases=sheet5_data[129][2:]
sheet5_Essex_cases=sheet5_data[130][2:]
sheet5_Gloucestershire_cases=sheet5_data[131][2:]
sheet5_Hampshire_cases=sheet5_data[132][2:]
sheet5_Hertfordshire_cases=sheet5_data[133][2:]
sheet5_Kent_cases=sheet5_data[134][2:]
sheet5_Lancashire_cases=sheet5_data[135][2:]
sheet5_Leicestershire_cases=sheet5_data[136][2:]
sheet5_Lincolnshire_cases=sheet5_data[137][2:]
sheet5_Norfolk_cases=sheet5_data[138][2:]
sheet5_Northamptonshire_cases=sheet5_data[139][2:]
sheet5_NYorkshire_cases=sheet5_data[140][2:]
sheet5_Nottinghamshire_cases=sheet5_data[141][2:]
sheet5_Oxfordshire_cases=sheet5_data[142][2:]
sheet5_Somerset_cases=sheet5_data[143][2:]
sheet5_Staffordshire_cases=sheet5_data[144][2:]
sheet5_Suffolk_cases=sheet5_data[145][2:]
sheet5_Surrey_cases=sheet5_data[146][2:]
sheet5_Warwickshire_cases=sheet5_data[147][2:]
sheet5_WSussex_cases=sheet5_data[148][2:]
sheet5_Worcestershire_cases=sheet5_data[149][2:]
sheet5_England_cases=sheet5_data[150][2:]

sheet5_data_req=[]
sheet5_data_req.append(['Area','Date','Cases'])

for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hartlepool',sheet5_date[i],sheet5_Hartlepool_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Middlesbrough',sheet5_date[i],sheet5_Middlesbrough_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Redcar and Cleveland',sheet5_date[i],sheet5_RC_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Stockton-on-Tees',sheet5_date[i],sheet5_Tees_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Darlington',sheet5_date[i],sheet5_Darlington_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Halton',sheet5_date[i],sheet5_Halton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Warrington',sheet5_date[i],sheet5_Warrington_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Blackburn with Darwen',sheet5_date[i],sheet5_Blackburn_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Blackpool',sheet5_date[i],sheet5_Blackpool_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Kingston upon Hull, City of',sheet5_date[i],sheet5_Kingston_Hull_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['East Riding of Yorkshire',sheet5_date[i],sheet5_ERY_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['North East Lincolnshire',sheet5_date[i],sheet5_NEL_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['North Lincolnshire',sheet5_date[i],sheet5_NL_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['York',sheet5_date[i],sheet5_York_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Derby',sheet5_date[i],sheet5_Derby_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Leicester',sheet5_date[i],sheet5_Leicester_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Rutland',sheet5_date[i],sheet5_Rutland_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Nottingham',sheet5_date[i],sheet5_Nottingham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Herefordshire, County of',sheet5_date[i],sheet5_Herefordshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Telford and Wrekin',sheet5_date[i],sheet5_Telford_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Stoke-on-Trent',sheet5_date[i],sheet5_Stoke_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bath and North East Somerset',sheet5_date[i],sheet5_Bath_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bristol, City of',sheet5_date[i],sheet5_Bristol_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['North Somerset',sheet5_date[i],sheet5_Nsomerset_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['South Gloucestershire',sheet5_date[i],sheet5_SGloucestershire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Plymouth',sheet5_date[i],sheet5_Plymouth_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Torbay',sheet5_date[i],sheet5_Torbay_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Swindon',sheet5_date[i],sheet5_Swindon_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Peterborough',sheet5_date[i],sheet5_Peterborough_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Luton',sheet5_date[i],sheet5_Luton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Southend-on-Sea',sheet5_date[i],sheet5_Southend_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Thurrock',sheet5_date[i],sheet5_Thurrock_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Medway',sheet5_date[i],sheet5_Medway_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bracknell Forest',sheet5_date[i],sheet5_Bracknell_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['West Berkshire',sheet5_date[i],sheet5_WBerkshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Reading',sheet5_date[i],sheet5_Reading_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Slough',sheet5_date[i],sheet5_Slough_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Windsor and Maidenhead',sheet5_date[i],sheet5_Windsor_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wokingham',sheet5_date[i],sheet5_Wokingham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Milton Keynes',sheet5_date[i],sheet5_Milton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Brighton and Hove',sheet5_date[i],sheet5_Brighton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Portsmouth',sheet5_date[i],sheet5_Portsmouth_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Southampton',sheet5_date[i],sheet5_Southampton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Isle of Wight',sheet5_date[i],sheet5_Isle_Wight_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['County Durham',sheet5_date[i],sheet5_Durham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Cheshire East',sheet5_date[i],sheet5_CheshireE_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Cheshire West and Chester',sheet5_date[i],sheet5_CheshireW_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Shropshire',sheet5_date[i],sheet5_Shropshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Cornwall and Isles of Scilly',sheet5_date[i],sheet5_Cornwall_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wiltshire',sheet5_date[i],sheet5_Wiltshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bedford',sheet5_date[i],sheet5_Bedford_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Central Bedfordshire',sheet5_date[i],sheet5_CBedfordshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Northumberland',sheet5_date[i],sheet5_Northumberland_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bournemouth, Christchurch and Poole',sheet5_date[i],sheet5_Bournemouth_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Dorset',sheet5_date[i],sheet5_Dorset_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bolton',sheet5_date[i],sheet5_Bolton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bury',sheet5_date[i],sheet5_Bury_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Manchester',sheet5_date[i],sheet5_Manchester_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Oldham',sheet5_date[i],sheet5_Oldham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Rochdale',sheet5_date[i],sheet5_Rochdale_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Salford',sheet5_date[i],sheet5_Salford_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Stockport',sheet5_date[i],sheet5_Stockport_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Tameside',sheet5_date[i],sheet5_Tameside_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Trafford',sheet5_date[i],sheet5_Trafford_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wigan',sheet5_date[i],sheet5_Wigan_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Knowsley',sheet5_date[i],sheet5_Knowsley_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Liverpool',sheet5_date[i],sheet5_Liverpool_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['St. Helens',sheet5_date[i],sheet5_Helens_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Sefton',sheet5_date[i],sheet5_Sefton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wirral',sheet5_date[i],sheet5_Wirral_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Barnsley',sheet5_date[i],sheet5_Barnsley_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Doncaster',sheet5_date[i],sheet5_Doncaster_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Rotherham',sheet5_date[i],sheet5_Rotherham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Sheffield',sheet5_date[i],sheet5_Sheffield_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Newcastle upon Tyne',sheet5_date[i],sheet5_Newcastle_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['North Tyneside',sheet5_date[i],sheet5_Ntyneside_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['South Tyneside',sheet5_date[i],sheet5_STyneside_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Sunderland',sheet5_date[i],sheet5_Sunderland_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Birmingham',sheet5_date[i],sheet5_Birmingham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Coventry',sheet5_date[i],sheet5_Coventry_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Dudley',sheet5_date[i],sheet5_Dudley_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Sandwell',sheet5_date[i],sheet5_Sandwell_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Solihull',sheet5_date[i],sheet5_Solihull_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Walsall',sheet5_date[i],sheet5_Walsall_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wolverhampton',sheet5_date[i],sheet5_Wolverhampton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bradford',sheet5_date[i],sheet5_Bradford_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Calderdale',sheet5_date[i],sheet5_Calderdale_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Kirklees',sheet5_date[i],sheet5_Kirklees_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Leeds',sheet5_date[i],sheet5_Leeds_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wakefield',sheet5_date[i],sheet5_Wakefield_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Gateshead',sheet5_date[i],sheet5_Gateshead_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Barking and Dagenham',sheet5_date[i],sheet5_Barking_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Barnet',sheet5_date[i],sheet5_Barnet_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bexley',sheet5_date[i],sheet5_Bexley_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Brent',sheet5_date[i],sheet5_Brent_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Bromley',sheet5_date[i],sheet5_Bromley_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Camden',sheet5_date[i],sheet5_Camden_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Croydon',sheet5_date[i],sheet5_Croydon_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Ealing',sheet5_date[i],sheet5_Ealing_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Enfield',sheet5_date[i],sheet5_Enfield_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Greenwich',sheet5_date[i],sheet5_Greenwich_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hackney and City of London',sheet5_date[i],sheet5_Hackney_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hammersmith and Fulham',sheet5_date[i],sheet5_Hammersmith_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Haringey',sheet5_date[i],sheet5_Haringey_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Harrow',sheet5_date[i],sheet5_Harrow_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Havering',sheet5_date[i],sheet5_Havering_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hillingdon',sheet5_date[i],sheet5_Hillingdon_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hounslow',sheet5_date[i],sheet5_Hounslow_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Islington',sheet5_date[i],sheet5_Islington_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Kensington and Chelsea',sheet5_date[i],sheet5_Kensington_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Kingston upon Thames',sheet5_date[i],sheet5_Kingston_Thames_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Lambeth',sheet5_date[i],sheet5_Lambeth_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Lewisham',sheet5_date[i],sheet5_Lewisham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Merton',sheet5_date[i],sheet5_Merton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Newham',sheet5_date[i],sheet5_Newham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Redbridge',sheet5_date[i],sheet5_Redbridge_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Richmond upon Thames',sheet5_date[i],sheet5_Richmond_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Southwark',sheet5_date[i],sheet5_Southwark_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Sutton',sheet5_date[i],sheet5_Sutton_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Tower Hamlets',sheet5_date[i],sheet5_Tower_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Waltham Forest',sheet5_date[i],sheet5_Waltham_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Wandsworth',sheet5_date[i],sheet5_Wandsworth_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Westminster',sheet5_date[i],sheet5_Westminster_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Buckinghamshire',sheet5_date[i],sheet5_Buckinghamshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Cambridgeshire',sheet5_date[i],sheet5_Cambridgeshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Cumbria',sheet5_date[i],sheet5_Cumbria_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Derbyshire',sheet5_date[i],sheet5_Derbyshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Devon',sheet5_date[i],sheet5_Devon_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['East Sussex',sheet5_date[i],sheet5_ESussex_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Essex',sheet5_date[i],sheet5_Essex_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Gloucestershire',sheet5_date[i],sheet5_Gloucestershire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hampshire',sheet5_date[i],sheet5_Hampshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Hertfordshire',sheet5_date[i],sheet5_Hertfordshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Kent',sheet5_date[i],sheet5_Kent_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Lancashire',sheet5_date[i],sheet5_Lancashire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Leicestershire',sheet5_date[i],sheet5_Leicestershire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Lincolnshire',sheet5_date[i],sheet5_Lincolnshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Norfolk',sheet5_date[i],sheet5_Norfolk_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Northamptonshire',sheet5_date[i],sheet5_Northamptonshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['North Yorkshire',sheet5_date[i],sheet5_NYorkshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Nottinghamshire',sheet5_date[i],sheet5_Nottinghamshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Oxfordshire',sheet5_date[i],sheet5_Oxfordshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Somerset',sheet5_date[i],sheet5_Somerset_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Staffordshire',sheet5_date[i],sheet5_Staffordshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Suffolk',sheet5_date[i],sheet5_Suffolk_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Surrey',sheet5_date[i],sheet5_Surrey_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Warwickshire',sheet5_date[i],sheet5_Warwickshire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['West Sussex',sheet5_date[i],sheet5_WSussex_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['Worcestershire',sheet5_date[i],sheet5_Worcestershire_cases[i]])
for i in range(len(sheet5_date)):
    sheet5_data_req.append(['England' ,sheet5_date[i],sheet5_England_cases[i]])

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
    
    
#sheet 1 excel
UK_Cases_COVID_UK_data_path = os.path.join(current_directory, r'UK_Cases_COVID_UK.xls')
if os.path.exists(UK_Cases_COVID_UK_data_path):
    os.remove(UK_Cases_COVID_UK_data_path) 
else:
    pass    
workbook=xlsxwriter.Workbook(r'UK_Cases_COVID_UK.xls') #EDIT 
worksheet1=workbook.add_worksheet()
for i in range(0,len(sheet1_data)):
    for j in range(0,3):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet1.write(cor,str(sheet1_data[i][j]))
        
workbook.close()

#sheet 2 excel
UK_Deaths_COVID_UK_data_path = os.path.join(current_directory, r'UK_Deaths_COVID_UK.xls')
if os.path.exists(UK_Deaths_COVID_UK_data_path):
    os.remove(UK_Deaths_COVID_UK_data_path) 
else:
    pass   
workbook2=xlsxwriter.Workbook(r'UK_Deaths_COVID_UK.xls') #EDIT 
worksheet2=workbook2.add_worksheet()
for i in range(0,len(l)):
    for j in range(0,len(l[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet2.write(cor,str(l[i][j]))
        
workbook2.close()

#sheet 3 excel
Countries_COVID_UK_data_path = os.path.join(current_directory, r'Countries_COVID_UK.xls')
if os.path.exists(Countries_COVID_UK_data_path):
    os.remove(Countries_COVID_UK_data_path) 
else:
    pass   
workbook3=xlsxwriter.Workbook(r'Countries_COVID_UK.xls') #EDIT 
worksheet3=workbook3.add_worksheet()
for i in range(0,len(sheet3_data_req)):
    for j in range(0,len(sheet3_data_req[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet3.write(cor,str(sheet3_data_req[i][j]))

workbook3.close()

#sheet 5 excel
UTLAs_COVID_UK_data_path = os.path.join(current_directory, r'UTLAs_COVID_UK.xls')
if os.path.exists(UTLAs_COVID_UK_data_path):
    os.remove(UTLAs_COVID_UK_data_path) 
else:
    pass  
workbook5=xlsxwriter.Workbook(r'UTLAs_COVID_UK.xls') #EDIT 
worksheet5=workbook5.add_worksheet()
for i in range(0,len(sheet5_data_req)):
    for j in range(0,len(sheet5_data_req[i])):
        r,c=i+1,j+1
        cor=chr(65+j)+str(r)
        worksheet5.write(cor,str(sheet5_data_req[i][j]))
        
workbook5.close()


#sheet 6 excel
Recovered_Patients_COVID_UK_data_path = os.path.join(current_directory, r'Recovered_Patients_COVID_UK.xls')
if os.path.exists(Recovered_Patients_COVID_UK_data_path):
    os.remove(Recovered_Patients_COVID_UK_data_path) 
else:
    pass  
workbook6=xlsxwriter.Workbook(r'Recovered_Patients_COVID_UK.xls') #EDIT 
worksheet6=workbook6.add_worksheet()
worksheet6.write('A1','Date')
worksheet6.write('B1','Cumulative Counts')
for i in range(1,len(sheet6_data)):
    a1='A'+str(i+1)
    worksheet6.write(a1,str(sheet6_data[i][0]))
    b1='B'+str(i+1)
    worksheet6.write(b1,str(sheet6_data[i][1]))
    
workbook6.close()

ACCESS_KEY = 'I'
SECRET_KEY = '1'

#sheet 1 s3
boto3.client('s3', aws_access_key_id=ACCESS_KEY,aws_secret_access_key=SECRET_KEY).upload_file("UK_Cases_COVID_UK.xls", "analyst-adhoc", "COVID_DASHBOARD_EMEA/UK_DATA_UNION_HEALTH/UK_Cases_COVID_UK.xlsx")

#sheet 2 s3
boto3.client('s3', aws_access_key_id=ACCESS_KEY,aws_secret_access_key=SECRET_KEY).upload_file("UK_Deaths_COVID_UK.xls", "analyst-adhoc", "COVID_DASHBOARD_EMEA/UK_DATA_UNION_HEALTH/UK_Deaths_COVID_UK.xlsx")

#sheet 3 s3
boto3.client('s3', aws_access_key_id=ACCESS_KEY,aws_secret_access_key=SECRET_KEY).upload_file("Countries_COVID_UK.xls", "analyst-adhoc", "COVID_DASHBOARD_EMEA/UK_DATA_UNION_HEALTH/Countries_COVID_UK.xlsx")

#sheet 5 s3
boto3.client('s3', aws_access_key_id=ACCESS_KEY,aws_secret_access_key=SECRET_KEY).upload_file("UTLAs_COVID_UK.xls", "analyst-adhoc", "COVID_DASHBOARD_EMEA/UK_DATA_UNION_HEALTH/UTLAs_COVID_UK.xlsx")

#sheet 6 s3
boto3.client('s3', aws_access_key_id=ACCESS_KEY,aws_secret_access_key=SECRET_KEY).upload_file("Recovered_Patients_COVID_UK.xls", "analyst-adhoc", "COVID_DASHBOARD_EMEA/UK_DATA_UNION_HEALTH/Recovered_Patients_COVID_UK.xlsx")