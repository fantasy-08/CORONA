import requests 
from bs4 import BeautifulSoup 
import csv 
import time
def convert_to_number(strr):
    num=''
    for i in range(len(strr)):
        if(strr[i]==','):continue
        else:num+=strr[i]
    return int(num)
f=open('C:\\Users\\eshaa\\Desktop\\CORONA\\Data.txt')
content=f.readlines()
f.close()
#OLD DATA COLLECTION
old_total=''
for i in content[0][14:]:
    if(i=='\n'):
        break
    else:
        old_total+=i
old_deaths=''
for i in content[1][14:]:
    if(i=='\n'):
        break
    else:
        old_deaths+=i
old_rec=''
for i in content[2][11:]:
    if(i=='\n'):
        break
    else:
        old_rec+=i
old_rec=int(old_rec)
old_total=int(old_total)
old_deaths=int(old_deaths)
old_time =content[3]
URL = "https://www.worldometers.info/coronavirus/"
r = requests.get(URL) 
  
soup = BeautifulSoup(r.content, 'html5lib') 
  
quotes=[]
  
table = soup.findAll('div', attrs = {'class':'maincounter-number'}) 
print("TIME")
print (time.asctime( time.localtime(time.time()) )+'     '+'since '+old_time)
print()
total_cases=(table[0].span.text)
total_cases=convert_to_number(total_cases)
siz='+'
if(total_cases<old_total):siz=''
print('TOTAL CASES- '+str(total_cases)+'   '+siz,((total_cases-old_total)))

total_deaths=(table[1].span.text)
total_deaths=convert_to_number(total_deaths)
siz='+'
if(total_cases<old_total):siz=''
print('TOTAL Deaths- '+str(total_deaths)+'   '+siz,((total_deaths-old_deaths)))

total_rec=(table[2].span.text)
total_rec=convert_to_number(total_rec)
siz='+'
print('Recovered- '+str(total_rec)+'   '+siz,((total_rec-old_rec)))

f=open('C:\\Users\\eshaa\\Desktop\\CORONA\\Data.txt','w')
f.write('Total Cases : '+str(total_cases)+'\n')
f.write('Total Deaths : '+str(total_deaths)+'\n')
f.write('Recovered : '+str(total_rec)+'\n')
time1=(time.asctime( time.localtime(time.time()) ))
f.write(time1)
f.close()
