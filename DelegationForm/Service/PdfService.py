from Base import Constant
import os
import pdfrw
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics,ttfonts
from Model import Delegation

class PdfService():

    def Work(self, xlsList):
        for delegationDict in xlsList:
            self.SetOverlay(delegationDict, self.InitCanvas())
            self.MergePdf(Constant.GetS89CH(), Constant.GetXlsFolder() + "\\temp.pdf", Constant.GetOutputFolder() + "\\" + delegationDict[Delegation.DcitName()])

    def InitCanvas(self):
        pdfmetrics.registerFont (ttfonts.TTFont ('chinese', Constant.GetFont_msjhbd()))  # 註冊字型
        if os.path.exists(Constant.GetXlsFolder() + "\\temp.pdf"):
            os.remove(Constant.GetXlsFolder() + "\\temp.pdf")
        cv = canvas.Canvas(Constant.GetXlsFolder() + "\\temp.pdf")
        cv.setFont ('chinese', 10)
        return cv

    def SetOverlay(self, delegationDict, cv):
        cv.drawString(115, 100, delegationDict[Delegation.DcitName()])
        '''cv.drawString(115, 650, delegationDict[Delegation.DictDate()])
        cv.drawString(115, 700, delegationDict[Delegation.DictDelegate()])
        cv.drawString(115, 750, delegationDict[Delegation.DictAssistant()])
        cv.drawString(115, 800, delegationDict[Delegation.DictClass()])'''
        cv.save()

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