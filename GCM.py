def openfile(filename):
	try:
		proc = subprocess.Popen(filename, shell=True, stdout=subprocess.PIPE)
		proc.wait()
		print("The programme exits with return code of:")
		print(proc.returncode)
		return
	except:
		print("An error occured.")

print("Running application")
print("Importing openpyxl")
from openpyxl import Workbook, load_workbook as lwb
print("Importing subprocess")
import subprocess
print("Importing sys")
import sys
print("Importing PyQt5")
from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtCore import QDate, QTime
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
class Applic(QWidget):
	def __init__(self):
		super().__init__()
		self.__initUI__()

	def __initUI__(self):
		self.setGeometry(300,300,300,220)
		self.setWindowTitle("Yosh")
		self.show()
		self.task = [Tasks()]
		print(self.task[0].toString())
if __name__ == '__main__':
	app = QApplication(sys.argv)
	ex = Applic()
	app.exec_()
#openfile("output.xlsm")