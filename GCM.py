print("Running application")
print("Importing openpyxl")
from openpyxl import Workbook, load_workbook
print("Importing subprocess")
import subprocess
print("Importing sys")
print("Importing PyQt5")
from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtCore import QDate, QTime
class FileIO():
	def openfile(filename):
		try:
			proc = subprocess.Popen(filename, shell=True, stdout=subprocess.PIPE)
			proc.wait()
			print("The programme exits with return code of:")
			print(proc.returncode)
		except:
			print("An error occured.")
	def createBlank():
		sb = load_workbook('Resource/file.xlsm', read_only = False, keep_vba = True)
		ss = sb['Sheet1']
		count = 2
		#print(ss['A' + str(count)].value)
		while(ss['A' + str(count)].value != None):
			ss['A' + str(count)] = None
			ss['B' + str(count)] = None
			ss['C' + str(count)] = None
			count += 1
			if (count == 100):
				print("Ended potentionally endless loop")
		sb.save('output.xlsm')
	def insertValues(values):
		sb = load_workbook('output.xlsm', read_only = False, keep_vba = True)
		ss = sb['Sheet1']
		for i in range(len(values)):
#			print(values[i].returnDateObj().toString('dd/MM/yyyy'))
			ss['A' + str(i + 2)] = values[i].returnTaskName()
			ss['B' + str(i + 2)] = values[i].returnDateObj().toString('dd/MM/yyyy')
			ss['C' + str(i + 2)] = values[i].returnDayDuration()
		sb.save('output.xlsm')
class Tasks:
	taskCount = 0
	def __init__(self, dat = QDate(2001, 1, 1), ddu = 0, tms = QTime(0,0,0,0), tme = QTime(23, 59, 0,0), tan = "", pri = 0):
		Tasks.taskCount += 1
		self.date = dat
		self.dayDuration = ddu
		self.timeStart = tms
		self.timeEnd = tme
		self.priority = pri
		if tan == "":
			self.taskName = "Task " + str(Tasks.taskCount)
		else:
			self.taskName = tan
	def toString(self):
		toReturn = self.taskName + ": " + self.date.toString() + " - " + self.date.addDays(self.dayDuration).toString()
		if (self.priority == 0):
			toReturn += ", Low Priority"
		elif (self.priority == 1):
			toReturn += ", Medium Priority"
		else:
			toReturn += ", High Priority"
		return toReturn
	def returnDateObj(self):
		return self.date
	def returnTaskName(self):
		return self.taskName
	def returnDayDuration(self):
		return self.dayDuration
class Applic(QWidget):
	def __init__(self):
		super().__init__()
		self.tasksList = []
		self.__initUI__()

	def __initUI__(self):
		self.setGeometry(300,300,300,220)
		self.setWindowTitle("Yosh")
		self.show()

	def createTask(self, dat = None, ddu = None, tms = None, tme = None, tan = None, pri = None):
		if (dat == None or ddu == None or tms == None or tme == None or tan == None or pri == None):
			print("Default value for the parameters is used.")
			self.tasksList.append(Tasks())
		else:
			self.tasksList.append(Tasks(dat, ddu, tms, tme, tan, pri))
		print("Task created: " + self.tasksList[len(self.tasksList) - 1].toString())
	
	def deleteTask(self, taskNumber = None):
		if (taskNumber == None):
			print("Deleting the last task.")
			taskNumber = len(self.tasksList)
		else:
			print("Deleting task number " + str(taskNumber))
		try:
			self.tasksList.remove(self.tasksList[taskNumber-1])
		except:
			print("Deletion not successful")
	def commitWork(self):
		FileIO.createBlank()
		FileIO.insertValues(self.tasksList)
if __name__ == '__main__':
	app = QApplication([])
	ex = Applic()
	ex.createTask()
	ex.createTask()
	ex.createTask()
	ex.commitWork()
	app.exec_()
#FileIO.openfile("output.xlsm")