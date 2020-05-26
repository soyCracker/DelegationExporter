from Base import Config
from Base import Constant
from Service.ExcelService import ExcelService
from Service.PdfService import PdfService
import os

def Main():
    Prepare()    
    WritePdf(ReadFromExcel())

def Prepare():
    print(Config.GetVersion())
    InitDir()

def ReadFromExcel():
    excelService = ExcelService()
    return excelService.ReadDelegation()

def WritePdf(xlsList):
    pdfService = PdfService()
    pdfService.Work(xlsList)


def InitDir():
    if os.path.exists(Constant.GetXlsFolder())==False:
        os.makedirs(Constant.GetXlsFolder())
    if os.path.exists(Constant.GetOutputFolder())==False:
        os.makedirs(Constant.GetOutputFolder())

Main()