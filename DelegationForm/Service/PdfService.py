from Base import Constant
import os
import pdfrw
from Util import TimeUtil 
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics,ttfonts
from Model import Delegation

class PdfService():

    def __init__(self):
        self.outputPath=""

    def Work(self, xlsList):
        self.InitFolder()
        for delegationDict in xlsList:
            self.SetOverlay(delegationDict, self.InitCanvas())
            self.MergePdf(Constant.GetS89CH(), Constant.GetXlsFolder() + "\\temp.pdf", self.outputPath + "\\" + delegationDict[Delegation.DcitName()] + ".pdf")

    def InitFolder(self):
        self.outputPath = Constant.GetOutputFolder() + "\\" + TimeUtil.GetCurrentTime()
        os.makedirs(self.outputPath)  

    def InitCanvas(self):
        pdfmetrics.registerFont (ttfonts.TTFont ('chinese', Constant.GetFont_msjhbd()))  # 註冊字型
        if os.path.exists(Constant.GetXlsFolder() + "\\temp.pdf"):
            os.remove(Constant.GetXlsFolder() + "\\temp.pdf")
        cv = canvas.Canvas(Constant.GetXlsFolder() + "\\temp.pdf")
        cv.setFont ('chinese', 10)
        return cv

    def SetOverlay(self, delegationDict, cv):
        if delegationDict[Delegation.DcitName()]!=None:
            cv.drawString(40, 280, delegationDict[Delegation.DcitName()])
        if delegationDict[Delegation.DictDate()]!=None and delegationDict[Delegation.DictDelegate()]!=None:
            cv.drawString(40, 230, delegationDict[Delegation.DictDate()] + " - " + delegationDict[Delegation.DictDelegate()])
        if delegationDict[Delegation.DictAssistant()]!=None:
            cv.drawString(40, 255, delegationDict[Delegation.DictAssistant()])
        self.SetClass(delegationDict, cv)
        cv.save()

    def SetClass(self, delegationDict, cv):
        if delegationDict[Delegation.DictClass()]==1:
            cv.drawString(15, 120, "V")
        else:
            cv.drawString(15, 105, "V")

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

'''pdfService = PdfService()
pdfService.OpenPdf(Constant.GetXlsFolder() + "\\temp")'''