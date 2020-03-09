
import requests 
from bs4 import BeautifulSoup 
import csv 
  
URL = "https://www.worldometers.info/coronavirus/"
r = requests.get(URL) 
  
soup = BeautifulSoup(r.content, 'html5lib') 
  
quotes=[]
  
table = soup.find('div', attrs = {'class':'maincounter-number'}) 
print(table.span.text)
