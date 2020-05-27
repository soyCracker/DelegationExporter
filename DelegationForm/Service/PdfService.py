from Base import Constant
from Util import TimeUtil 
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics,ttfonts
from Model import Delegation
from Model import DelegationType
import os
import pdfrw

class PdfService():

    def __init__(self):
        self.outputPath = Constant.GetOutputFolder() + "\\" + TimeUtil.GetCurrentTime()
        self.tempPdf = Constant.GetXlsFolder() + "\\temp.pdf"
        self.jpStr = "JP"
        self.description = "傳道與生活聚會委派通知單-"
        self.jpDescription = "日語-"

    def Work(self, xlsList):
        print("處理中...")
        self.InitFolder()      
        for delegationDict in xlsList:
            print(".......")
            self.SetOverlay(delegationDict, self.InitCanvas())
            self.MergePdf(self.WhichS89(delegationDict), self.tempPdf, self.GetOutputPdfName(delegationDict))
        print("完成")

    def InitFolder(self):
        os.makedirs(self.outputPath)  

    def InitCanvas(self):
        pdfmetrics.registerFont (ttfonts.TTFont ('chinese', Constant.GetFont_msjhbd()))  # 註冊字型
        if os.path.exists(self.tempPdf):
            os.remove(self.tempPdf)
        cv = canvas.Canvas(self.tempPdf)
        cv.setFont ('chinese', 10)
        return cv

    def SetOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DcitName()].find(self.jpStr)>=0:
            self.SetJpOverlay(delegationDict, cv)
        else:
            self.SetChineseOverlay(delegationDict, cv)

    def SetChineseOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DcitName()]!=None:
            cv.drawString(40, 280, delegationDict[Delegation.DcitName()].replace(self.jpStr, ""))
        if delegationDict[Delegation.DictDate()]!=None and delegationDict[Delegation.DictDelegate()]!=None:
            cv.drawString(40, 230, delegationDict[Delegation.DictDate()] + " - " + delegationDict[Delegation.DictDelegate()])
        if delegationDict[Delegation.DictAssistant()]!=None:
            cv.drawString(40, 255, delegationDict[Delegation.DictAssistant()])
        self.SetChineseDelegationOverlay(delegationDict, cv)
        self.SetChineseClassOverlay(delegationDict, cv)
        cv.save()

    def SetChineseClassOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DictClass()]==1:
            cv.drawString(15, 120, "V")
        else:
            cv.drawString(15, 105, "V")

    def SetChineseDelegationOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DictDelegate()].find(DelegationType.Reading())>=0:
            cv.drawString(15, 190, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.InitialCall())>=0:
            cv.drawString(15, 175, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.FirstRV())>=0:
            cv.drawString(15, 160, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.SecondRV())>=0:
            cv.drawString(15, 150, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.BibleStudy())>=0:
            cv.drawString(110, 190, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.BibleStudy2())>=0:
            cv.drawString(110, 190, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.Talk())>=0:
            cv.drawString(110, 175, "V")
        #續訪判斷放最後面
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.FirstRV2())>=0:
            cv.drawString(15, 160, "V")
        else:
            print("SetDelegationOverlay else")

    def SetJpOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DcitName()]!=None:
            cv.drawString(40, 263, delegationDict[Delegation.DcitName()].replace(self.jpStr, ""))
        if delegationDict[Delegation.DictDate()]!=None and delegationDict[Delegation.DictDelegate()]!=None:
            cv.drawString(40, 220, delegationDict[Delegation.DictDate()] + " - " + delegationDict[Delegation.DictDelegate()])
        if delegationDict[Delegation.DictAssistant()]!=None:
            cv.drawString(40, 240, delegationDict[Delegation.DictAssistant()])
        self.SetJpDelegationOverlay(delegationDict, cv)
        self.SetJpClassOverlay(delegationDict, cv)
        cv.save()

    def SetJpClassOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DictClass()]==1:
            cv.drawString(19, 97, "V")
        else:
            cv.drawString(19, 85, "V")

    def SetJpDelegationOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DictDelegate()].find(DelegationType.Reading())>=0:
            cv.drawString(19, 158, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.InitialCall())>=0:
            cv.drawString(19, 147, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.FirstRV())>=0:
            cv.drawString(19, 134, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.SecondRV())>=0:
            cv.drawString(19, 122, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.BibleStudy())>=0:
            cv.drawString(135, 147, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.BibleStudy2())>=0:
            cv.drawString(135, 147, "V")
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.Talk())>=0:
            cv.drawString(135, 134, "V")
        #續訪判斷放最後面
        elif delegationDict[Delegation.DictDelegate()].find(DelegationType.FirstRV2())>=0:
            cv.drawString(19, 134, "V")
        else:
            print("SetDelegationOverlay else")

    def MergePdf(self, form_pdf, overlay_pdf, output):
        """
        Merge the specified fillable form PDF with the 
        overlay PDF and save the output
        """
        form = pdfrw.PdfReader(form_pdf)
        olay = pdfrw.PdfReader(overlay_pdf)
    
        for form_page, overlay_page in zip(form.pages, olay.pages):
            merge_obj = pdfrw.PageMerge()
            overlay = merge_obj.add(overlay_page)[0]
            pdfrw.PageMerge(form_page).add(overlay).render()
        
        writer = pdfrw.PdfWriter()
        writer.write(output, form)

    def WhichS89(self, delegationDict):
        if delegationDict[Delegation.DcitName()].find(self.jpStr)>=0:
            return Constant.GetS89J()
        else:
            return Constant.GetS89CH()

    def GetOutputPdfName(self, delegationDict):
        if delegationDict[Delegation.DcitName()].find(self.jpStr)>=0:
            return self.outputPath + "\\" + self.description + self.jpDescription + delegationDict[Delegation.DcitName()].replace(self.jpStr, "") + ".pdf"
        else:
            return self.outputPath + "\\" + self.description + delegationDict[Delegation.DcitName()] + ".pdf"
