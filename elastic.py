import json,re
import xlwt 
from xlwt import Workbook

#es = open("index_alert.json")
#data = json.load(es)
#filedata = open("ihub.json")
wb = Workbook()
sheet1 = wb.add_sheet('Analysis')
sheet1.write(0, 0, 'Alert Unique Key') 
sheet1.write(0, 1, 'IHUB Status') 
sheet1.write(0, 2, 'Alert Store Status') 


def extract(text):
 regex = '.*\[(.*?)\].*'
 matches = re.search(regex, text)
 line = matches.group(1)
 return line

def parseAlerts():
 for hit in data['hits']:
  alertMessage = hit['_source']['alertMessage']
  extractedValue = extract(alertMessage)
  alertKey = "["+extractedValue+"]"
  status = searchInIHub(alertKey)
  print(alertKey,status)

def searchInIHub(alertKey,filename):
  with open(filename, 'r') as read_obj:
   for line in read_obj:
    if alertKey in line:
     print(alertKey,"Found in",filename)
     return True
   print(alertKey,"not in",filename)
   return False

try:
 i = 1;
 while(True):
  alertKey = '[BBB222 '+str(i)+']'
  sheet1.write(i,0 , alertKey)
  ihubstatus = searchInIHub(alertKey,"ihub.json")
  alertstorestatus = searchInIHub(alertKey,"index_alert.json")
  sheet1.write(i,1 , ihubstatus)
  sheet1.write(i,2 , alertstorestatus)
  if(i%10):
   wb.save('Analysis.xls') 
  i = i + 1
  if(i==4433):
   wb.save('Analysis.xls') 
   break
except Exception as ex:
  print (ex)
 
