from datetime import datetime
import xlrd, time
from xlwt import Workbook

loc = "d:\\work\\sanjay\\du2.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

wbout = Workbook()
sheetout = wbout.add_sheet('Output')

def format():
 try:
    rownum = 1
    for row in range(1, sheet.nrows):
     start = sheet.cell_value(rownum,6)
     sheetout.write(rownum,1,start)
     if(type(start) is str):
      end = sheet.cell_value(rownum,7)
      convert(start,end)
     rownum = rownum + 1
     if (rownum % 10):
      wbout.save('output.xls')
     if (rownum == 350):
      wbout.save('output.xls')
 except Exception as ex:
  print(ex)
	
def getcurrentdatetime():
 today = datetime.datetime.now()
 format = "%Y-%m-%dT%H:%M:%S.%f"
 today = today.strftime(format)
 print(today)

def convert(starttime,endtime):
 datetimeFormat = "%Y-%m-%dT%H:%M:%S.%f"
 start = datetime.strptime(starttime, "%Y-%m-%dT%H:%M:%S.%f")
 end = datetime.strptime(endtime, "%Y-%m-%dT%H:%M:%S.%f")
 time_dif = datetime.strptime(endtime, datetimeFormat) - datetime.strptime(starttime,datetimeFormat)
 print(start, end, time_dif.total_seconds()*1000)
format()
