from xlrd import open_workbook
#from xlrd import open_workbook
from collections import OrderedDict
from openDefectGUI import *

#import sys
#from PyQt5.QtWidgets import (QMainWindow, QTextEdit, QAction, QFileDialog, QApplication)
#from PyQt5.QtGui import QIcon

import xlwt
import sys





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
        print('Header',self.header)
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
            print (index)
            value = self.rowConvert(index)
            print('Value',value)
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

                tempDict['Defect ID']=key
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
        self.sheet = None
       # self.sheet = self.sumWb.add_sheet('Sheet1', cell_overwrite_ok=True)

    def createSheet(self,name='Sheet1'):
        self.sheet = self.sumWb.add_sheet(name,cell_overwrite_ok=True)
        return True

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

if __name__ == "__main__":

    if len(sys.argv) < 2:
        print('Fehler')




    else:
        print (len(sys.argv))
       # test = wrapper(sys.argv)
        defectFile = sys.argv[1]
        meetingFile = sys.argv[2]
        resultFile = sys.argv[3]
        print('File',defectFile,meetingFile)


    defect = XLSConvert(defectFile)
    meeting = XLSConvert(meetingFile)
  #  print('Sheets', dc.getSheets)
    defect.selectSheet('Sheet1')
    defect.setHeader(0)
    defect.convertSheet()

    meeting.selectSheet('Sheet1')
    meeting.setHeader(0)
    meeting.convertSheet()
    print(meeting.getData())
   #print('Row',dc.getRow(0))
   # dc.test()
   # defect.getHeader(0)
  #  print (defect.convertSheet())
    combine = compare(defect.getData(),meeting.getData())

#    print(combine.getData())

    write = writeXls(combine.getData())
    write.createSheet('Sheet1')
    write.createTable()
    write.writeFile(resultFile)
  #  print (combine.getData())


   # comp = compare(defect.convertXLS(),meeting.convertXLS())
   # result = comp.test()
  #  print (result)

 #   write = writeXls(result)
#   # .convertXLS()