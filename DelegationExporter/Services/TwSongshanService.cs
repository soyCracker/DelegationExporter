using DelegationExporter.Base;
using DelegationExporter.Except;
using DelegationExporter.Interface;
using DelegationExporter.Model;
using DelegationExporter.Util;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace DelegationExporter.Services
{
    public class TwSongshanService : IDelegationService
    {
        private string blockDate = "";
        private string lastDelagation = "";
        private string lastDate = "";
        private int sameDelegationCount = 1;

        public List<T> ReadDelegation<T>(string filePath)
        {
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Open);
                IWorkbook workbook;
                try
                {
                    workbook = new XSSFWorkbook(fs);
                }
                catch (Exception)
                {
                    workbook = new HSSFWorkbook(fs);
                }
                ISheet sheet = workbook.GetSheetAt(0);
                List<S89Xlsx> s89List = new List<S89Xlsx>();
                for(int j=1;j<=2;j++)
                {
                    for (int i = 4; i <= sheet.LastRowNum; i++)
                    {
                        SetXlsxList(sheet, i, s89List, j);
                    }
                }
                workbook.Close();
                fs.Close();
                return s89List as List<T>;
            }
            catch (IOException)
            {
                throw;
            }
        }

        public void WriteDelegation<T>(List<T> delegationList, string destFolder)
        {
            List<S89Xlsx> list = delegationList as List<S89Xlsx>;
            foreach (S89Xlsx delegation in list)
            {
                Console.WriteLine("-------------------------------\n");
                string s89 = Config.S89CH;
                string description = "傳道與生活聚會委派通知單-";
                if (delegation.Name.Contains("JP"))
                {
                    s89 = Config.S89J;
                    description = description + "日語-";
                    delegation.Name = delegation.Name.Replace("JP", "");
                }
                WriteDelegationInLanguage(s89, delegation, destFolder, description);               
            }
        }

        private void WriteDelegationInLanguage(string s89, S89Xlsx delegation, string destFolder, string description)
        {
            using (FileStream fs = new FileStream(Config.FILE_FOLDER + "//" + s89, FileMode.Open))
            {    
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(PdfUtil.FileNameExistAddR(destFolder + "//" + description + delegation.Name+".pdf")));
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                SetPdfField(fields, delegation, pdfDoc);
                pdfDoc.Close();
            }
        }

        private void SetXlsxList(ISheet sheet, int rowNum, List<S89Xlsx> s89List, int classInt)
        {
            S89Xlsx s89Model = new S89Xlsx();
            try
            {
                s89Model.Name = sheet.GetRow(rowNum).GetCell(2 + (classInt - 1) * 2).ToString();
                if( string.IsNullOrEmpty(s89Model.Name) )
                {
                    return;
                }
            }
            catch (Exception)
            {
                return;
            }
            try
            {
                s89Model.Assistant = sheet.GetRow(rowNum).GetCell(3 + (classInt-1)*2 ).ToString();
            }
            catch (Exception)
            {
                s89Model.Assistant = "";
            }
            s89Model.Delegation = StringUtil.GetChinesePrintAble( sheet.GetRow(rowNum).GetCell(1).ToString() );
            s89Model.Class = classInt.ToString();
            try
            {
                s89Model.Date = sheet.GetRow(rowNum).GetCell(0).ToString();
                if ( string.IsNullOrEmpty(s89Model.Date) )
                {
                    s89Model.Date = blockDate;
                }
                else
                {
                    blockDate = s89Model.Date;
                }             
            }
            catch(Exception)
            {
                s89Model.Date = blockDate;
            }
            s89List.Add(s89Model);
        }

        private void SetPdfFieldDelegationCell(IDictionary<string, PdfFormField> fields, S89Xlsx delegation, PdfDocument pdfDoc)
        {
            Console.WriteLine("委派:" + delegation.Delegation + "\n");
            if(delegation.Delegation.Contains(S89chDelegationName.Reading))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Reading);
                lastDelagation = S89chDelegationName.Reading;
            }
            else if(delegation.Delegation.Contains(S89chDelegationName.InitialCall))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.InitialCall);
                lastDelagation = S89chDelegationName.InitialCall;
            }
            else if (delegation.Delegation.Contains(S89chDelegationName.FirstRV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FirstRV);
                lastDelagation = S89chDelegationName.FirstRV;
            }
            else if (delegation.Delegation.Contains(S89chDelegationName.SecondRV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.SecondRV);
                lastDelagation = S89chDelegationName.SecondRV;
            }
            else if (delegation.Delegation.Contains(S89chDelegationName.BibleStudy))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BibleStudy);
                lastDelagation = S89chDelegationName.BibleStudy;
            }
            else if (delegation.Delegation.Contains(S89chDelegationName.BibleStudy2))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BibleStudy);
                lastDelagation = S89chDelegationName.BibleStudy2;
            }
            else if (delegation.Delegation.Contains(S89chDelegationName.Talk))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Talk);
                lastDelagation = S89chDelegationName.Talk;
            }
            //續訪擺在最後避免與第二次續訪衝突
            else if (delegation.Delegation.Contains(S89chDelegationName.FirstRV2))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FirstRV);
                lastDelagation = S89chDelegationName.FirstRV2;
            }
            else
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Other);
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.OtherText, delegation.Delegation.Split('(')[0]);
            }
        }

        private void SetPdfFieldClass(IDictionary<string, PdfFormField> fields, S89Xlsx delegation, PdfDocument pdfDoc)
        {
            Console.WriteLine("班別:" + delegation.Class + "\n");
            switch (delegation.Class)
            {
                case "1":
                    PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Class1);
                    break;
                case "2":
                    PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Class2);
                    break;
                default:
                    Console.WriteLine("班別填錯了吧\n");
                    break;
            }
        }

        private void SetPdfFieldDelegationName(IDictionary<string, PdfFormField> fields, S89Xlsx delegation, PdfDocument pdfDoc)
        {
            if(delegation.Date.Equals(lastDate) && delegation.Delegation.Contains(lastDelagation))
            {
                sameDelegationCount++;
                Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Delegation.Replace(lastDelagation, lastDelagation + sameDelegationCount) + "\n");
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" + delegation.Delegation.Replace(lastDelagation, lastDelagation + sameDelegationCount));
            }
            else
            {
                sameDelegationCount=1;
                Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Delegation + "\n");
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" + delegation.Delegation);
            }
            lastDate = delegation.Date;
        }

        private void SetPdfField(IDictionary<string, PdfFormField> fields, S89Xlsx delegation, PdfDocument pdfDoc)
        {
            if (PdfUtil.IsMsjhbdExist())
            {
                //我她X的，SetFont要在SetValue之前
                Console.WriteLine("學生:" + delegation.Name + "\n");
                PdfUtil.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.Name, delegation.Name);

                Console.WriteLine("助手:" + delegation.Assistant + "\n");
                PdfUtil.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.Ass, delegation.Assistant);

                SetPdfFieldDelegationName(fields, delegation, pdfDoc);

                /*Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Delegation + "\n");
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" + delegation.Delegation);*/

                SetPdfFieldDelegationCell(fields, delegation, pdfDoc);

                SetPdfFieldClass(fields, delegation, pdfDoc);
            }
            else
            {
                throw new NoFontException();
            }
        }

        public string BeforePrepareAndGetTempXlsx()
        {
            List<string> fileList = FileUtil.GetFileList(Config.FILE_FOLDER);
            string tempXls = "";
            foreach(string file in fileList)
            {
                if(file.ToLowerInvariant().EndsWith(".xls") || file.ToLowerInvariant().EndsWith(".xlsx"))
                {
                    if (File.Exists(file + Config.TEMP_NAME))
                    {
                        File.Delete(file + Config.TEMP_NAME);
                    }
                    FileUtil.CopyTemp(file, file + Config.TEMP_NAME);
                    tempXls = file + Config.TEMP_NAME;
                }
            }
            Console.WriteLine("BeforePrepareAndGetTempXlsx() tempXls:" + tempXls + "\n");
            return tempXls;         
        }
    }
}
