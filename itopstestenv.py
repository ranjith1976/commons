import xlrd, time, datetime,sys
import smtplib,ssl
from re import search
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
count89 = 0
count1000 = 0

#loc="d:\\work\\files\\excel\\tests.xlsx"
#loc="d:\\work\\testsetup\\SolarWinds.xlsx"
loc="/home/mluser/ITOps/loadtest/SolarWinds.xlsx"
wb = xlrd.open_workbook(loc)
#gmail_users = ['pttest2107@gmail.com','pttest2108@gmail.com','pttest2110@gmail.com','pttest2111@gmail.com']
#gmail_pwds = ['lemkkpubuelochez','wbruvrdafcuvrozw','vexvfgjsoiytwmph','xntdzxdwxlvdispt']
#to = ['autops.demo@gmail.com']
#to = ['ranjith.mnair@ust-global.com']
to = ['SmartOpsQAinbound@ust-global.com']
gmail_users = ['sajumust@gmail.com','smartops.ustglobal@gmail.com','smartops.privateuser@gmail.com','smartops.privateuser1@gmail.com','smartops.privateuser2@gmail.com','smartops.privateuser3@gmail.com','smartops.privateuser4@gmail.com','smartops.privateuser5@gmail.com','smartops.privateuser6@gmail.com','smartops.privateuser7@gmail.com','pttest2107@gmail.com','pttest2108@gmail.com','pttest2110@gmail.com','smartops.privateuser8@gmail.com']
gmail_pwds = ["ebojculpglcppane","sfyfxezhiskzuzse","iznzzmruhznaurcl","dxcmpsapwvzvttfw","gegubbocaysfydfk","jufjapqviylumpdr","agaslvhekopjddps","nqazexcgocwtierx","sotajwfaooprdmmv","vlvdsbjgbnlmdrui","lemkkpubuelochez","wbruvrdafcuvrozw","vexvfgjsoiytwmph","enkrtupebpmxvudu"]
gmail_status={}

def check_gmailstatus(gmail):
  global gmail_status
  if(gmail in gmail_status):
   gmailcounter = gmail_status[gmail]
   gmailcounter = gmailcounter + 1
   gmail_status[gmail] = gmailcounter
  else:
   gmail_status[gmail] = 1

def readexcel(startfrom=1,simulation=True):
 try:

  print("Read excel", simulation)
  sheet = wb.sheet_by_index(0)
  rownum = startfrom
  alerts = list()
  endcount = sheet.nrows
  for row in range(startfrom, sheet.nrows):
     alert = generateAlert(sheet,rownum)
     alerts.append(alert)
     rownum = rownum + 1
  response= prepareemail(alerts,simulation)
  if response is None:
      return False
  if rownum < endcount:
   return True
  else:
   return False
 except Exception as ex:
  timenow = getcurrentdatetime(0)
  log("FAILED","EMAIL",ex)

def generateAlert(sheet,rownum):
    alert = {}
    alertCorrelationKey = sheet.cell_value(rownum, 0)
    alert['alertCorrelationKey'] = int(alertCorrelationKey)
    alertDate = sheet.cell_value(rownum, 1)
    alert['alertDate'] = alertDate
    alertSubject = sheet.cell_value(rownum, 2)
    alert['alertSubject'] = alertSubject
    alertMessage = sheet.cell_value(rownum, 3)
    alert['alertMessage'] = alertMessage
    alertName = sheet.cell_value(rownum, 4)
    alert['alertName'] = alertName
    alertSeverity = sheet.cell_value(rownum, 5)
    alert['alertSeverity'] = alertSeverity
    alertTime = sheet.cell_value(rownum, 6)
    alert['alertTime'] = alertTime
    alertDuration = sheet.cell_value(rownum, 7)
    alert['alertDuration'] = int(alertDuration)
    alertServiceName = sheet.cell_value(rownum, 8)
    alert['alertServiceName'] = alertServiceName
    alertNodeName = sheet.cell_value(rownum, 9)
    alert['alertNodeName'] = alertNodeName
    alertNodeIpAddr = sheet.cell_value(rownum, 10)
    alert['alertNodeIpAddr'] = alertNodeIpAddr
    alertObjectType = sheet.cell_value(rownum, 11)
    alert['alertObjectType'] = alertObjectType
    alertObjectName = sheet.cell_value(rownum, 12)
    alert['alertObjectName'] = alertObjectName
    alertObjectStatus = sheet.cell_value(rownum, 13)
    alert['alertObjectStatus'] = alertObjectStatus
    alertUrl = sheet.cell_value(rownum, 14)
    alert['alertUrl'] = alertUrl
    return alert

def generatemsg(alert):
    if(alert!=None):
        timenow = getcurrentdatetime()
        message = """\
         <html>
           <head></head>
           <body>
             Alert Message: [TEMP123 %s] %s</br> 
             Alert Name: %s</br>
             Alert Severity: %s</br>
             Alert Time: %s,</br>
             Service Name: %s</br>
             Node Name: %s</br>
             Node IP Address: %s</br>
             Object Type: %s</br>
             Object Name: %s</br>
             Object Status: %s </br>
             Alert Details Url: %s
           </body>
         </html>
         """ % (alert['alertCorrelationKey'],alert['alertMessage'],alert['alertName'], alert['alertSeverity'], timenow, alert['alertServiceName'],
                alert['alertNodeName'], alert['alertNodeIpAddr'], alert['alertObjectType'], alert['alertObjectName'],
                alert['alertObjectStatus'], alert['alertUrl'])

        return message
    else:
        return None

def prepareemail(alerts,simulate):
 global count89
 global count1000
 emailcounter=0
 i = 0
 smtpserver = getsmtp()
 while(i < len(alerts)):
   alert = alerts[i]
   message = generatemsg(alert)
   if(message!=None):
       duration = alert['alertDuration']
       log("Sleeping", "Duration "+str(duration),'')
       time.sleep(duration)
       status = sendemail(message,alert['alertSubject'],alert['alertCorrelationKey'],smtpserver)
       if (status == "MAX_LIMIT"):
        print(gmail_status)
        print("Max limit reached",emailcounter)
        try:
         print("Current Email",gmail_users[emailcounter])
         emailcounter = emailcounter + 1
         if(emailcounter == len(gmail_users)):
            print("Resetting to 0")
            emailcounter = 0
         print("New Email",gmail_users[emailcounter])
        except Exception as ex:
            log("EXCEPTION", "Prepare Email Max Limit", ex)
        if(emailcounter == len(gmail_users)):
            emailcounter = 0
        smtpserver = getsmtp(emailcounter)
       else:
        i = i + 1
   else: 
       return None
def getsmtp(emailcounter=0):
    try:
        log("LOGGING IN", "SMTP", "")
        context = ssl.create_default_context()
        smtpserver = smtplib.SMTP("smtp.gmail.com", 587)
        smtpserver.starttls()
        smtpserver.ehlo
        smtpserver.login(gmail_users[emailcounter], gmail_pwds[emailcounter])
        loggedinuser = gmail_users[emailcounter]
        log("LOGGED IN", "SMTP", "")
        return smtpserver
    except Exception as ex:
        log("EXCEPTION", "SMTP", ex)
        return None

def sendemail(message,subject,correlationkey,smtpserver):
 global count89
 global count1000
 global gmail_status
 try:
  subject = subject + " [Key: " +str(correlationkey)+"]"
  msg = MIMEMultipart()
  msg['To'] = to[0]
  msg['Subject'] = subject
  msg.attach(MIMEText(message,"html"))
  if (smtpserver!=None):
    try:
     smtpserver.sendmail(smtpserver.user,msg['To'], msg.as_string())
     count89 = count89 + 1
     count1000 = count1000 + 1
     check_gmailstatus(smtpserver.user)
     if count89 == 89:
         count89 = 0
         log('MAX Limit Exception ' + smtpserver.user, "EMAIL", "")
         return "MAX_LIMIT"
     if count1000 == 1000:
         log('MAX Limit Error ' + smtpserver.user, "EMAIL", "")
         return "MAX_LIMIT"
         count1000 = 0
    except smtplib.SMTPDataError as error:
        print("Error",error)
        if error.smtp_code == 550 or error.smtp_code == 421:
            timenow = getcurrentdatetime(0)
            log('MAX Limit Error '+smtpserver.user, "EMAIL", "")
            if (search("Daily user sending quota exceeded", str(error.smtp_error)) or search("Try again later, closing connection", str(error.smtp_error))):
                return "MAX_LIMIT"
    except Exception as ex:
        print("Ex",ex)
        strval = str(ex)
        timenow = getcurrentdatetime(0)
        log('MAX Limit Exception ' + smtpserver.user, "EMAIL", "")
        if (search("Try again later, closing connection", strval)):
            return "MAX_LIMIT"
        return "ERROR"
    timenow = getcurrentdatetime(0)
    log('SUCCESS '+ subject, "EMAIL"," ",correlationkey)
    return "SUCCESS"
  else:
    timenow = getcurrentdatetime(0)
    log('FAILED SMTP None ' + subject, "EMAIL", "Error",correlationkey)
    return "ERROR"
 except Exception as ex:
  log('FAILED SMTP None ' + subject, "EMAIL", ex,correlationkey)
  return "ERROR"

def getcurrentdatetime(type=1):
 today = datetime.datetime.now()
 format = "%A, %B %d, %Y %H:%M"
 if(type==0):
     format = "%Y-%m-%d %H:%M:%S:%f"

 today = today.strftime(format)
 return today

def log(status,locationincode,exception,correlationKey=0 ):
    print(correlationKey, ',', status, ',', locationincode, ',', getcurrentdatetime(0), ',', exception)



def main(argv):
   startswith = 1
   alertsim = True
   try:
     stattswith = int(argv[0])
     if(argv[1] == "False"):
         alertsim = False
   except:
     print('First argument must be a number and second boolean - defaulting starts with to 1 and alert simulation True')
   if(readexcel(startswith,alertsim)):
      log('PROCESS COMPLETE ' , "", "",)
   else:
      log('PROCESS FAILED ', "", "", )
if __name__ == "__main__":
  try:
   main(sys.argv[1:])
  except:
      print("Error")
