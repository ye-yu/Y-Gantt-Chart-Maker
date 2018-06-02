#to read from file
from openpyxl import *
sb = load_workbook('test.xlsm', keep_vba  = True)
wb = Workbook()
ss = sb['Sheet1']
ws = wb.active
ws['B1'] = "Start Time"
ws['C1'] = "Duration"
count = 2
#print(ss['A' + str(count)].value)
while(ss['A' + str(count)].value != None):
	ws['A' + str(count)] = ss['A' + str(count)].value
	count += 1
	if (count == 100):
		print("Ended potentinally endless loop")
count = 2
while(ws['A' + str(count)].value != None):
	print(ws['A' + str(count)].value)
	count += 1