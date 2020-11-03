import datetime, math
import xlrd, time

loc="d:\\work\\files\\excel\\SolarWinds.xlsx"
wb = xlrd.open_workbook(loc)


def xldate_to_datetime(xldatetime):  # something like 43705.6158241088

 tempDate = datetime.datetime(1899, 12, 31)
 (days, portion) = math.modf(xldatetime)

 deltaDays = datetime.timedelta(days=days)
 # changing the variable name in the edit
 secs = int(24 * 60 * 60 * portion)
 detlaSeconds = datetime.timedelta(seconds=secs)
 TheTime = (tempDate + deltaDays + detlaSeconds)
 return TheTime
  #.strftime("%m/%d/%Y %H:%Ms")

def convert(date):
 try:
  #date = "Mon, 20 Apr 2020 21:12:05"
  format = "%A, %B %d, %Y %H:%M"
  format = "%m/%d/%Y %H:%M"
  #date = date.rstrip(',')

  date = xldate_to_datetime(date)

  return date
 except Exception as ex:
  print (ex)

def readexcel():
 try:
  sheet = wb.sheet_by_index(0)
  rownum = 1
  prevdate = None

  for row in range(1, sheet.nrows):
   if rownum != 1:
    prevdate = currentdate
   slno = sheet.cell_value(rownum, 0)

   currentdate = convert(sheet.cell_value(rownum, 6))
   #format = "%m/%d/%Y %H:%M"
   #newdate = currentdate.strftime(format)

   if prevdate != None:
    x = currentdate - prevdate
    print (int(slno),' ',int(x.seconds))
   else:
    print(print (int(slno),' ',0))
   rownum = rownum + 1
 except Exception as ex:
  print (ex)

def currentdatetime():
 today = datetime.datetime.now()
 format = "%A, %B %d, %Y %H:%M"
 today = today.strftime(format)

readexcel()
#currentdatetime()


