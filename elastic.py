import requests,json,re
import pyparsing  as pp
import json,re
from datetime import datetime
import xlwt 
from xlwt import Workbook

FILENAME = "Analysis_v5.xls"
MAX_ROW_COUNT=20
START_ROW = 1
ELASTICSEARCH_FILE = "ES_DDD444.json"

ihubdata = {}

def extractkey(text,left="[",right="]"):
 #print('[',text[text.index(left)+len(left):text.index(right)],']')
 word = pp.Word(pp.alphanums)
 s = text
 rule = pp.nestedExpr('[', ']')
 for match in rule.searchString(s):
     lst = match
     key = (match[0][0])
     if 'DDD444' in key:
        key = 'DDD444 '
        val = match[0][1]
        return '['+key+val+']'


def getIHUBJson():
 s = requests.Session()
 s.headers.update({'offline-token':"eyJhbGciOiJIUzI1NiIsInR5cCIgOiAiSldUIiwia2lkIiA6ICJhODdjYzgwOS02YTA1LTQyY2MtOTY3YS0zNjk3OGFjZGFkZTUifQ.eyJqdGkiOiJhYzc1MmQxNy0zNzI1LTQ3YTQtOTNlMy0wY2JkNzUyZjUzNWUiLCJleHAiOjAsIm5iZiI6MCwiaWF0IjoxNjAyNjU2MjY2LCJpc3MiOiJodHRwczovL3NtYXJ0b3BzLXB0MDEuZWFzdHVzLmNsb3VkYXBwLmF6dXJlLmNvbS9wYWFzL2l0b3BzL2tleWNsb2FrL2F1dGgvcmVhbG1zL3VzdGdsb2JhbCIsImF1ZCI6Imh0dHBzOi8vc21hcnRvcHMtcHQwMS5lYXN0dXMuY2xvdWRhcHAuYXp1cmUuY29tL3BhYXMvaXRvcHMva2V5Y2xvYWsvYXV0aC9yZWFsbXMvdXN0Z2xvYmFsIiwic3ViIjoiYjdiYTUxZWUtOTMxZi00NzA4LTgzYjYtZmRkYTA3ODk1MDQxIiwidHlwIjoiT2ZmbGluZSIsImF6cCI6InNtYXJ0b3BzLWZyb250ZW5kIiwiYXV0aF90aW1lIjowLCJzZXNzaW9uX3N0YXRlIjoiNTczZWI3YmEtZjE2NS00Y2IyLWFkNTUtOWZlZjJhMzJjOTY5IiwicmVhbG1fYWNjZXNzIjp7InJvbGVzIjpbIm9mZmxpbmVfYWNjZXNzIl19LCJzY29wZSI6Im9mZmxpbmVfYWNjZXNzIn0.Ctl2S3Xh7fBt0z1LH0uLIOAc_q9ExRIku5adl47SXd0"})
 s.headers.update({'Organization-key': '1'})
 s.headers.update({'Content-Type': 'application/json'})
 print("Fetching")
 response = s.get('https://smartops-pt01.eastus.cloudapp.azure.com/paas/itops/ihubservice/api/smartops/ihub/logs?interval=10000')
 js = {}
 if response.status_code ==200:
  js = response.json()
 else:
  print("Can't connect to ihub")
  exit(1) 

 for i in range(len(js)):
    outbound = js[i]["outbound"]
    inboundtime = js[i]['inboundTimeStamp']
    outboundtime = js[i]['outboundTimestamp']
    time = {}
    time['inboundtime'] = inboundtime
    time['outboundtime'] = outboundtime
    #print (outbound)
    outboundobj = json.loads(outbound)
    inputs = outboundobj['requestJson']['inputs']
    for inputindex in range(len(inputs)):
      requestMessage = inputs[inputindex]['requestMessage']
      if 'DDD444' in requestMessage:
       key = (extractkey(requestMessage))
       ihubdata[key] = time



def datediff(starttime,endtime):

 if starttime == "NONE" or endtime == "NONE":
  return 0
 datetimeFormat = "%Y-%m-%d %H:%M:%S.%f"
 start = datetime.strptime(starttime, "%Y-%m-%d %H:%M:%S.%f")
 end = datetime.strptime(endtime, "%Y-%m-%d %H:%M:%S.%f")
 time_dif = datetime.strptime(endtime, datetimeFormat) - datetime.strptime(starttime,datetimeFormat)
 time_diff_in_millines = time_dif.total_seconds() * 1000
 return time_dif.total_seconds()

def extract(text,left,right): 
 line = text[text.index(left)+len(left):text.index(right)]
 line = "["+line+"]"
 return line


def convertformat(time,format="NORMAL"):
 if (time == "NONE"):
  return "NONE"
 if (format=="SEC"):
  datetimeFormat = "%Y-%m-%d %H:%M:%S:%f"
  timeasis = datetime.strptime(time, datetimeFormat)
  returndate = timeasis.strftime("%Y-%m-%d %H:%M:%S.%f")
 elif(format=="UTC"):
  datetimeFormat = "%Y-%m-%dT%H:%M:%S.%f%z"
  timeasis = datetime.strptime(time, datetimeFormat)
  returndate = timeasis.strftime("%Y-%m-%d %H:%M:%S.%f")
 else:
  datetimeFormat = "%Y-%m-%dT%H:%M:%S.%f"
  timeasis = datetime.strptime(time, datetimeFormat)
  returndate = timeasis.strftime("%Y-%m-%d %H:%M:%S.%f")

 return returndate

es = open(ELASTICSEARCH_FILE)
alertData = {}
data = {}
with open(ELASTICSEARCH_FILE) as f:
 data = json.load(f)

length = len(data['hits']['hits'])
i = 0
for i in range(i,length):
 keystr = None
 alertObj = data['hits']['hits'][i]
 if "_source" in alertObj:
  if "alertMessage" in alertObj["_source"]:
   alertMessage = alertObj["_source"]["alertMessage"]
   keystr = extract(alertMessage,"[","]")
  else:
   alertMessage = "NONE"

  if "source" in alertObj["_source"]:
   source = alertObj["_source"]["source"]
  else:
   source = "NONE"

  if "modifiedTime" in alertObj["_source"]:
   modifiedTimeReal = alertObj["_source"]["modifiedTime"]
   modifiedTime = convertformat(modifiedTimeReal)
  else:
   modifiedTime = "NONE"

  if "requestReceivedTime" in alertObj["_source"]:
   requestReceivedTimeReal = alertObj["_source"]["requestReceivedTime"]
   requestReceivedTime = convertformat(requestReceivedTimeReal,"UTC")
  else:
   requestReceivedTime = "NONE"

  if "createdTime" in alertObj["_source"]:
   createdTimeReal = alertObj["_source"]["createdTime"]
   createdTime = convertformat(createdTimeReal)
  else:
   createdTime = "NONE"

  if "alertTime" in alertObj["_source"]:
   alertTimeReal = alertObj["_source"]["alertTime"]
   alertTime = convertformat(alertTimeReal)
  else:
   alertTime = "NONE"

  if "ticketStatus" in alertObj["_source"]:
   ticketStatus = alertObj["_source"]["ticketNumber"]
  else:
   ticketStatus = "NONE"

  ticketedTime = "NONE"
  if "clusterInfo" in alertObj["_source"]:
   if "status"  in alertObj["_source"]["clusterInfo"]:
    status = alertObj["_source"]["clusterInfo"]["status"]
    if status == "open":
     ticketedTime = "NONE"
    elif status == "ticketed":
     if "statusTime" in alertObj["_source"]["clusterInfo"]:
      ticketedTime = alertObj["_source"]["clusterInfo"]["statusTime"]
    else:
     if "statusHistory" in alertObj["_source"]["clusterInfo"]:
      statusHistoryItems = alertObj["_source"]["clusterInfo"]['statusHistory']
      for statusHistoryItem in statusHistoryItems:
       if statusHistoryItem["status"] == "ticketed":
        ticketedTime = statusHistoryItem["statusTime"]
  if(ticketedTime is not None and ticketedTime!="NONE"):
   ticketedTime = convertformat(ticketedTime)
  result = alertMessage[alertMessage.index("[DDD444") + len("[DDD444"):alertMessage.index("]")]
  diff = 0
  alert = {} 
  if keystr in alertData:
   alertExisting = alertData[keystr]
   existingModifiedTime = alertExisting["modifiedTime"]
   diff = datediff(modifiedTime,existingModifiedTime)
   #print (keystr, modifiedTime, existingModifiedTime,diff)
   if diff >0:
     alert = alertExisting
  if diff <= 0:
   alert["alertMessage"] = alertMessage
   alert["alertTime"] = alertTime
   alert["createdTime"] = createdTime
   alert["modifiedTime"] = modifiedTime
   alert["requestReceivedTime"] = requestReceivedTime
   alert["source"] = source
   alert["ticketStatus"] = ticketStatus
   alert["ticketedTime"] = ticketedTime

  alertData[keystr]=alert

  #if "[DDD444 4]" in alertData:
  #    print(keystr,"Found")
  #else:
  #    print(keystr,"Not Found")
 else:
  print("error in elastic search index - _Source not found")
  exit(1)
getIHUBJson()
filedata = open("iHub_log_DDD444.json")
wb = Workbook()
sheet1 = wb.add_sheet(FILENAME)
sheet1.write(0, 0, 'Alert Unique Key') 
sheet1.write(0, 1, 'IHUB Status') 
sheet1.write(0, 2, 'Alert Store Status')
sheet1.write(0, 3, 'IHUB Count')
sheet1.write(0, 4, 'Alert Count')
sheet1.write(0, 5, 'Alert Time (AT)')
sheet1.write(0, 6, 'Received Time (RT)')
sheet1.write(0, 7, 'IHUB Inbound Time (IIT)')
sheet1.write(0, 8, 'IHUB Outbound Time (IOT)')
sheet1.write(0, 9, 'Created Time (CT)')
sheet1.write(0, 10, 'Modified Time (MT)')
sheet1.write(0, 11, 'Ticketed Time (TT)')

sheet1.write(0, 12, 'Source')
sheet1.write(0, 13, 'Ticket Status')
sheet1.write(0, 14, 'IIT - AR')
sheet1.write(0, 15, 'IOT - IIT')
sheet1.write(0, 16, 'IOT - AC')
sheet1.write(0, 17, 'TC - AC')
sheet1.write(0, 18, 'TC - AR (Over all)')

ihubfilecontent = None




def parseAlerts():
 for hit in data['hits']:
  alertMessage = hit['_source']['alertMessage']
  extractedValue = extract(alertMessage)
  alertKey = "["+extractedValue+"]"
  status = searchInIHub(alertKey)

def searchInIHub(alertKey):
 ihubStatus = False
 countIHub = ihubfilecontent.count(alertKey)
 if(countIHub>0):
  ihubStatus = True
 return countIHub,ihubStatus



def searchInAlert(alertKey):
 with open(ELASTICSEARCH_FILE, 'r') as read_obj:
  alertStatus = False
  countAlert = 0
  for line in read_obj:
   if alertKey in line:
    countAlert = countAlert + 1
 if(countAlert>0):
   alertStatus = True
 return countAlert,alertStatus


print(len(alertData))

try:
 with open("iHub_log_DDD444.json", 'r') as read_obj:
  for line in read_obj:
   ihubfilecontent = line

 i = START_ROW;
 while(True):
  alertKey = '[DDD444 '+str(i)+']'
  sheet1.write(i,0 , alertKey)
  print("Searching in ihub")
  countIHub,ihubstatus = searchInIHub(alertKey)
  print("Searching in aletstore")
  countAlert,alertstatus = searchInAlert(alertKey)
  print("Done")

  sheet1.write(i,1, ihubstatus)
  sheet1.write(i,2, alertstatus)
  sheet1.write(i,3, countIHub)
  sheet1.write(i,4, countAlert)
  print("Retriving alertkey from alert data")
  if alertKey in alertData:
   alertFromMap = alertData[alertKey]
   print("Retrieved")
   alertReceived = alertFromMap["requestReceivedTime"]
   alertCreated = alertFromMap["createdTime"]
   alertModified =alertFromMap["modifiedTime"]
   alertTime = alertFromMap["alertTime"]
   ticketCreatedTime = alertFromMap["ticketedTime"]

   alertSource = alertFromMap["source"]
   alertTicketStatus = alertFromMap["ticketStatus"]
   inboundtime = "NONE"
   outboundtime = "NONE"
   time = ihubdata[alertKey]   
   inboundtime = time['inboundtime']
   inboundtime = convertformat(inboundtime,"SEC")
   outboundtime = time['outboundtime']
   outboundtime = convertformat(outboundtime,"SEC")

   ar_iit = 0
   print(alertKey,alertReceived,inboundtime,outboundtime,alertCreated,alertModified,ticketCreatedTime)
   if(alertReceived != "NONE" and inboundtime != "NONE") :
    ar_iit = datediff(alertReceived,inboundtime)
    print("AR - IIT",alertReceived,inboundtime,ar_iit)
    if (ar_iit<0):
        ar_iit=0

   iit_iot=0
   if(inboundtime != "NONE" and outboundtime != "NONE") :
    iit_iot = datediff(inboundtime,outboundtime)
    print("IIT - IOT",inboundtime,outboundtime,iit_iot)
    if(iit_iot<0):
        iit_iot=0

   iot_ac = 0
   if (alertCreated != "NONE" and outboundtime != "NONE"):
    iot_ac = datediff(outboundtime,alertCreated)
    print("IOT - AC",outboundtime,alertCreated,iot_ac)
    if(iot_ac<0):
        iot_ac=0


   ac_tc = 0

   if (alertCreated != "NONE" and ticketCreatedTime != "NONE"):
    ac_tc = datediff(alertCreated,ticketCreatedTime)
    print("AC - TC", alertCreated, ticketCreatedTime,ac_tc)
    if(ac_tc<0):
        ac_tc=0


   overall = 0
   if(alertReceived!="NONE" and ticketCreatedTime!="NONE"):
    overall = datediff(alertReceived,ticketCreatedTime)
    print("overall - AR-TC",ticketCreatedTime,alertReceived, overall)
    if(overall<0):
        overall=0

   sheet1.write(i,5, alertTime)
   sheet1.write(i,6, alertReceived)
   sheet1.write(i,7, inboundtime)
   sheet1.write(i,8, outboundtime)
   sheet1.write(i,9, alertCreated)
   sheet1.write(i,10, alertModified)
   sheet1.write(i,11, ticketCreatedTime)
   sheet1.write(i,12, alertSource)
   sheet1.write(i,13, alertTicketStatus)
   sheet1.write(i,14, ar_iit)
   sheet1.write(i,15, iit_iot)
   sheet1.write(i,16, iot_ac)
   sheet1.write(i, 17, ac_tc)
   sheet1.write(i,18, overall)
  if(i%10==0):
   wb.save(FILENAME)
   pass
  i = i + 1
  if(i==MAX_ROW_COUNT):
   wb.save(FILENAME)
   break
except Exception as ex:
  print (ex)
