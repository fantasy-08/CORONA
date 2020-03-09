import requests 
from bs4 import BeautifulSoup 
import csv 
def convert_to_number(strr):
    num=''
    for i in range(len(strr)):
        if(strr[i]==','):continue
        else:num+=strr[i]
    return int(num)
f=open('C:\\Users\\eshaa\\Desktop\\CORONA\\Data.txt')
content=f.readlines()
f.close()
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
old_total=int(old_total)
old_deaths=int(old_deaths)

URL = "https://www.worldometers.info/coronavirus/"
r = requests.get(URL) 
  
soup = BeautifulSoup(r.content, 'html5lib') 
  
quotes=[]
  
table = soup.findAll('div', attrs = {'class':'maincounter-number'}) 
print(len(table))

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


f=open('C:\\Users\\eshaa\\Desktop\\CORONA\\Data.txt','w')
f.write('Total Cases : '+str(total_cases)+'\n')
f.write('Total Deaths : '+str(total_deaths)+'\n')
f.close()
