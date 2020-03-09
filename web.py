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
old_total=int(old_total)
URL = "https://www.worldometers.info/coronavirus/"
r = requests.get(URL) 
  
soup = BeautifulSoup(r.content, 'html5lib') 
  
quotes=[]
  
table = soup.find('div', attrs = {'class':'maincounter-number'}) 
total_cases=(table.span.text)
total_cases=convert_to_number(total_cases)
print('TOTAL CASES- '+str(total_cases)+'   +'+str(abs(total_cases-old_total)))



f=open('C:\\Users\\eshaa\\Desktop\\CORONA\\Data.txt','w')
f.write('Total Cases : '+str(total_cases)+'\n')
f.close()