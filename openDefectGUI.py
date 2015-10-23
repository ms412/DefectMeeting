# Form implementation generated from reading ui file 'DefectMeeting.ui'
#
# Created by: PyQt5 UI code generator 5.4.1
#
# WARNING! All changes made in this file will be lost!

from xlrd import open_workbook
from collections import OrderedDict
import xlwt
import sys
import sip

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QTextEdit,QAction, QFileDialog, QApplication)
from PyQt5.QtGui import QIcon

class Ui_Form(QMainWindow):
    def __init__(self):

        self.defectFile = None
        self.meetingFile = None
        self.outFile = None

        super().__init__()
   # def __init__(self):
    #    QtGui.QWidget.__init__(self)
        self.setupUi(self)

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(358, 158)
        self.verticalLayoutWidget = QtWidgets.QWidget(Form)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(0, 10, 221, 131))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.lineEdit = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.lineEdit.setObjectName("lineEdit")
        self.verticalLayout.addWidget(self.lineEdit)
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.verticalLayout.addWidget(self.lineEdit_2)
        self.label_3 = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.verticalLayout.addWidget(self.lineEdit_3)
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(Form)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(250, 10, 81, 131))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton_3.setObjectName("pushButton_3")
        self.verticalLayout_2.addWidget(self.pushButton_3)
        self.pushButton = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout_2.addWidget(self.pushButton)
        self.pushButton_2 = QtWidgets.QPushButton(self.verticalLayoutWidget_2)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout_2.addWidget(self.pushButton_2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "Defect File"))
        self.label_2.setText(_translate("Form", "Meeting File"))
        self.label_3.setText(_translate("Form", "Output File"))
        self.pushButton_3.setText(_translate("Form", "Info"))
        self.pushButton.setText(_translate("Form", "Create"))
        self.pushButton_2.setText(_translate("Form", "Quit"))

        self.pushButton_2.clicked.connect(self.quitButton)
        self.pushButton.clicked.connect(self.createFile)

        #self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.returnPressed.connect(self.lineEditDefect)
        self.lineEdit_2.returnPressed.connect(self.lineEditMeeting)
        self.lineEdit_3.returnPressed.connect(self.lineEditOutfile)
       # openFile1.triggered.connect(self.openMeetingFile)


    def quitButton(self):
        print("Quit")
        return True

    def createFile(self):
        print('Test')
        if self.defectFile == None or self.meetingFile == None or self.outFile == None:
            print ('No file')
        else:
            defect = XLSConvert(self.defectFile)
            meeting = XLSConvert(self.meetingFile)
        #  print('Sheets', dc.getSheets)
            defect.selectSheet('Sheet1')
            defect.setHeader(0)
            defect.convertSheet()

            meeting.selectSheet('Sheet1')
            meeting.setHeader(0)
            meeting.convertSheet()
            print(meeting.getData())

            combine = compare(defect.getData(),meeting.getData())

            write = writeXls(combine.getData())
            write.createTable()
            write.writeFile(self.outFile)

        return True



    def lineEditDefect(self):
        path = QFileDialog.getOpenFileName(self, 'Open file', 'c:/')
        #print('Defect',path)
        self.defectFile = path[0]
        #print('ii',self.defectFile)
        self.lineEdit.setText(self.defectFile)
        return True

    def lineEditMeeting(self):
        path = QFileDialog.getOpenFileName(self, 'Open file', 'c:/')
        self.meetingFile = path[0]
        #print('Test',self.meetingFile)
        self.lineEdit_2.setText(self.meetingFile)
        return True

    def lineEditOutfile(self):
        path = QFileDialog.getSaveFileName(self, "Save file", "", ".xls")
        self.outFile = path[0]
        self.lineEdit_3.setText(self.outFile)
        return True


class XLSConvert():

    def __init__(self,file):
        self.wb = open_workbook(file)
        self.sheet = None

       # OrderedDict(test)
        self.row = OrderedDict()
        self.container = OrderedDict()


    def getSheets(self):
        return self.wb.sheet_names()

    def selectSheet(self,sheet):
        self.sheet = self.wb.sheet_by_name(sheet)
        return True

    def getRow(self,line):
        return self.sheet.row(line)

    def setHeader(self,rowId):
        self.header = self.sheet.row(rowId)
        return self.header

    def rowConvert(self,idx):
       # temp = OrderedDict()
       # for header, col in zip (self.sheet.row(0),self.sheet.row(idx)):
        for header, col in zip(self.header,self.getRow(idx)):

            #if 'Defect ID' in str(header.value):
             #   keyID = col.value
             #   print ('Test', header.value, int(keyID))
#                self.container['ID']=keyID
              #  temp['ID']=int(keyID)
           # else:
              #  self.container[header.value] = col.value
            self.row[str(header.value)]=(col.value)


    #    return int(keyID), temp
        return self.row

    def getValue(self, keyname):
        return self.row.get(str(keyname),None)


    def convertSheet(self):
   #     header = self.sheet.row(0)

        for index in range(self.sheet.nrows)[1:]:
          #  print (index)
            value = self.rowConvert(index)
            key = int(self.getValue('Defect ID'))
          #  key, value = self.rowConvert(index)
       #     print(key,value)
            self.container[key]=value.copy()

       # print('resutl',self.container)
        return self.container

    def getData(self):
        return self.container

class compare(object):

    def __init__(self,defect,meeting):

        self.defect = defect
        self.meeting = meeting

        self.combined = OrderedDict()


        self.test()

    def test(self):
        tempDict = OrderedDict()
        for key,value in self.defect.items():
         #   print(key,value)
          #  temp = value.get(key)
            state = value.get('Defect Status',None)
         #   print('Test',key,state)
            if 'Closed'  in state:
              #  print('Closed')
                None
            elif 'Fixed' in state:
              #  print('Fixed')
                None
            else:
               # print('ID',key,state)

                tempDict['ID']=key
                tempDict['Summary']=value.get('Summary')
                tempDict['Severity']=value.get('Severity')
                tempDict['Defect Status']=value.get('Defect Status')
                tempDict['Status Whiteboard']=value.get('Status Whiteboard')

                content = self.meeting.get(key, 'empty')
           #     print ('Content',content)
                if not 'empty' in content:
                   # print('Content',content)
                    for me_key,me_value in content.items():
                        if 'Status Whiteboard' in me_key:
                            None
                        elif 'Summary' in me_key:
                            None
                        elif 'Severity' in me_key:
                            None
                        elif 'ID' in me_key:
                            None
                        elif 'Defect Status' in me_key:
                            None
                        elif 'Priority' in me_key:
                            None
                        else:
                            tempDict[me_key]=me_value


               # print('HH',self.meeting.get(key,None))

                self.combined[key]=tempDict.copy()
                tempDict.clear()

        return self.combined
          #  elif 'Fixed' in state:

    def getData(self):
        return self.combined


class writeXls(object):
    def __init__(self, data):

        self.data = data
        self.sumWb = xlwt.Workbook()
      #  self.format = workbook.add_format({'bold': True, 'font_color': 'red'})
        self.sheet = self.sumWb.add_sheet('Sheet 1', cell_overwrite_ok=True)



    def createTable(self):
        row = 0
        col = 0
        for key, value in self.data.items():
            row = row +1
            col = 0
            for index,data in value.items():

                self.createIndex(col,index)
                self.sheet.write(row,col,data)
                col = col + 1

    def createIndex(self,col,index):
        print('index',col,index)
        self.sheet.write(0,col,index)
        return True

    def writeFile(self,filename):
        self.sumWb.save(filename)


if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Ui_Form()
    ex.show()
    sys.exit(app.exec_())

