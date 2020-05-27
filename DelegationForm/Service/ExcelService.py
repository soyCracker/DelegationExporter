from Base import Constant
from Model import Delegation
from Util import StringUtil
import xlwings
import os
import shutil

class ExcelService():

    def __init__(self):
        self.xlsList=[]
        self.delegateDate=""
        self.tempXls = Constant.GetXlsFolder() + "\\temp.xls"
        self.sheetName = "Sheet1"

    def InitAndGetExcelFile(self):
        if os.path.exists(self.tempXls):
            os.remove(self.tempXls)
        xls = self.FindFirstExcel()
        shutil.copy(Constant.GetXlsFolder() + "\\" + xls, self.tempXls)
        return self.tempXls

    def FindFirstExcel(self):
        allFileList = os.listdir(Constant.GetXlsFolder())
        for file in allFileList:
            if file.endswith("xls") or file.endswith("xlsx"):
                return file
        return ""

    def ReadDelegation(self):
        workbook = xlwings.Book(self.InitAndGetExcelFile())
        sheet = workbook.sheets[self.sheetName]
        self.ReadOneClass(sheet, 1)
        self.ReadOneClass(sheet, 2)
        return self.xlsList

    def ReadOneClass(self, sheet, classInt):
        nameCell = classInt * 2 + 1
        assistantCell = classInt * 2 + 2
        for row in range(5, sheet.api.UsedRange.Rows.count):
            if sheet.range(row, nameCell).value is None or sheet.range(row, nameCell).value == "":
                continue
            else:           
                if sheet.range(row, 1).value is not None and sheet.range(row, 1).value != "":
                    self.delegateDate = sheet.range(row, 1).value
                self.xlsList.append( self.SetDelegationDict(self.delegateDate, sheet.range(row, 2).value, sheet.range(row, nameCell).value, sheet.range(row, assistantCell).value, classInt) )               

    def SetDelegationDict(self, date, delegation, name, assistant, delegateClass):
        delegationDict = dict()
        delegationDict[Delegation.DictDate()] = StringUtil.ClearSpace(date)
        delegationDict[Delegation.DictDelegate()] = StringUtil.ClearSpace(delegation)
        delegationDict[Delegation.DcitName()] = StringUtil.ClearSpace(name)
        delegationDict[Delegation.DictAssistant()] = StringUtil.ClearSpace(assistant)  
        delegationDict[Delegation.DictClass()] = delegateClass
        print(delegationDict)
        return delegationDict
     
    def GetInfo(self, sheet):
        print("GetInfo:\n")
        print(sheet.api.UsedRange.Rows.count)
        print(sheet.api.UsedRange.Columns.count)
        print(sheet.used_range.shape)