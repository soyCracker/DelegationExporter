from Base import Config
from Base import Constant
from Service.ExcelService import ExcelService
import os

def Main():
    if Prepare()==True:    
        WritePdf(ReadFromExcel())

def Prepare():
    print(Config.GetVersion())
    if InitDir()==False:
        return False
    return True

def ReadFromExcel():
    excelService = ExcelService()
    return excelService.ReadDelegation()

def WritePdf(xlsList):
    print("")


def InitDir():
    if os.path.exists(Constant.GetXlsFolder())==False:
        os.makedirs(Constant.GetXlsFolder())
        return False
    if os.path.exists(Constant.GetOutputFolder())==False:
        os.makedirs(Constant.GetOutputFolder())
    return True

Main()