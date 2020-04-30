from Base import Constant
from Model import Delegation
import xlwings
import os
import shutil

class ExcelService():

    def __init__(self):
        self.xlsList=[]
        self.delegateDate=""

    def SheetName(self):
        return "Sheet1"

    def InitAndGetExcelFile(self):
        if os.path.exists(Constant.GetXlsFolder() + "\\temp.xls"):
            os.remove(Constant.GetXlsFolder() + "\\temp.xls")
        xls = self.FindFirstExcel()
        shutil.copy(Constant.GetXlsFolder() + "\\" + xls, Constant.GetXlsFolder() + "\\temp.xls")
        return Constant.GetXlsFolder() + "\\temp.xls"

    def FindFirstExcel(self):
        allFileList = os.listdir(Constant.GetXlsFolder())
        for file in allFileList:
            if file.endswith("xls") or file.endswith("xlsx"):
                return file
        return ""

    def ReadDelegation(self):
        workbook = xlwings.Book(self.InitAndGetExcelFile())
        sheet = workbook.sheets[self.SheetName()]
        #print(sht.range(5,3).value)
        self.ReadOneClass(sheet, 1)
        self.ReadOneClass(sheet, 2)
        return self.xlsList

    def ReadOneClass(self, sheet, classInt):
        nameCell = 2+(classInt-1)*2
        assistantCell = 3+(classInt-1)*2
        for row in range(5, sheet.api.UsedRange.Rows.count):
            if sheet.range(row, nameCell).value is None or sheet.range(row, nameCell).value == "":
                continue
            else:
                if sheet.range(row, 1).value is not None or sheet.range(row, 1).value != "":
                    self.delegateDate = sheet.range(row, 1).value
                self.xlsList.append( self.SetDelegationDict(self.delegateDate, sheet.range(row, 2).value, sheet.range(row, nameCell).value, sheet.range(row, assistantCell).value) )               

    def SetDelegationDict(self, date, delegation, name, assistant):
        delegationDict = dict()
        delegationDict[Delegation.DictDate] = date
        delegationDict[Delegation.DictDelegate] = delegation
        delegationDict[Delegation.DcitName] = name
        delegationDict[Delegation.DictAssistant] = assistant
        print(delegationDict)
        return delegationDict

    def GetInfo(self, sheet):
        print("GetInfo:\n")
        print(sheet.api.UsedRange.Rows.count)
        print(sheet.api.UsedRange.Columns.count)
        print(sheet.used_range.shape)


'''excelService =  ExcelService()
excelService.ReadDelegation()'''