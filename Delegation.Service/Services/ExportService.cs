using Delegation.Service.Except;
using Delegation.Service.Models;
using Delegation.Service.Utils;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Delegation.Service.Services
{
    public class ExportService
    {
        private string blockDate = "";
        private string lastDelagation = "";
        private string lastDate = "";
        private int sameDelegationCount = 1;
        private PDFService pdfService;

        public ExportService(string fontFolder)
        {
            pdfService = new PDFService(fontFolder);
        }

        public void Start(string outputFolder, string fileFolder, string tempName, string s89chFile,
            string s89jpFile, string descStr, string descJPStr, string JPFlagStr)
        {
            string tempXlsxFile = GetTempXlsx(fileFolder, tempName);
            List<DelegationVM> list = ReadDelegationFromAssignFile(tempXlsxFile);
            ExportDelegation(fileFolder, list, CreateDestFolder(outputFolder), s89chFile, s89jpFile, descStr, 
                descJPStr, JPFlagStr);
            File.Delete(tempXlsxFile);
        }

        public string GetTempXlsx(string fileFolder, string tempName)
        {
            string[] files = Directory.GetFiles(fileFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).ToArray();
            if (files.Length > 1)
            {
                Console.WriteLine("PrepareAndGetTempXlsx() MultipleXlsException\n");
                throw new MultipleXlsException();
            }
            string file = files[0];
            string tempXls = file + tempName;
            if (File.Exists(tempXls))
            {
                File.Delete(tempXls);
            }
            File.Copy(file, tempXls);
            Console.WriteLine("PrepareAndGetTempXlsx() tempXls:" + tempXls + "\n");
            return tempXls;
        }

        private string CreateDestFolder(string outputFolder)
        {
            if (!Directory.Exists(outputFolder + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(outputFolder + "//" + TimeUtil.GetTimeNow());
            }
            return outputFolder + "//" + TimeUtil.GetTimeNow();
        }

        public List<DelegationVM> ReadDelegationFromAssignFile(string filePath)
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
                List<DelegationVM> vmList = new List<DelegationVM>();
                for (int j = 1; j <= 2; j++)
                {
                    for (int i = 4; i <= sheet.LastRowNum; i++)
                    {
                        var temp = GetEachDelegation(sheet, i, j);
                        if(temp!=null)
                        {
                            vmList.Add(temp);
                        }                       
                    }
                }
                workbook.Close();
                fs.Close();
                return vmList;
            }
            catch (IOException)
            {
                throw;
            }
        }

        private DelegationVM GetEachDelegation(ISheet sheet, int rowNum, int classInt)
        {        
            try
            {
                DelegationVM vm = new DelegationVM();
                vm.Name = sheet.GetRow(rowNum).GetCell(2 + (classInt - 1) * 2).ToString();
                vm.Assistant = sheet.GetRow(rowNum).GetCell(3 + (classInt - 1) * 2).ToString();
                vm.Header = StringUtil.GetChinesePrintAble(sheet.GetRow(rowNum).GetCell(1).ToString());
                vm.Class = classInt.ToString();
                vm.Date = sheet.GetRow(rowNum).GetCell(0).ToString();
                
                if (string.IsNullOrEmpty(vm.Name))
                {                  
                    return null;
                }
                if(string.IsNullOrEmpty(vm.Date))
                {
                    vm.Date = blockDate;
                }
                else
                {
                    blockDate = vm.Date;
                }
                return vm;
            }
            catch (Exception)
            {
                return null;
            }      
        }

        public void ExportDelegation(string fileFolder, List<DelegationVM> delegationList, string destFolder, string s89chFile, 
            string s89jpFile, string descStr, string descJPStr, string JPFlagStr)
        {
            foreach (DelegationVM delegation in delegationList)
            {
                Console.WriteLine("-------------------------------\n");
                string s89 = s89chFile;
                string description = descStr;
                //識別日文委派
                if (delegation.Name.Contains(JPFlagStr))
                {
                    s89 = s89jpFile;
                    description = description + descJPStr;
                    //識別日文委派的字符替換掉
                    delegation.Name = delegation.Name.Replace(JPFlagStr, "");
                }
                WritePdf(fileFolder, s89, delegation, destFolder, description);
            }
        }

        private void WritePdf(string fileFolder, string s89, DelegationVM delegation, string destFolder, string description)
        {
            using (FileStream fs = new FileStream(Path.Combine(fileFolder, s89), FileMode.Open))
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs),
                    new PdfWriter(pdfService.FileNameExistAddR(Path.Combine(destFolder,
                    TimeUtil.CovertDateToFileNameStr(delegation.Date) + description + delegation.Name + ".pdf"))));
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                SetPdfField(fields, delegation, pdfDoc);
                pdfDoc.Close();
            }
        }

        private void SetPdfField(IDictionary<string, PdfFormField> fields, DelegationVM delegation, PdfDocument pdfDoc)
        {
            if (pdfService.IsMsjhbdExist())
            {
                //我她X的，SetFont要在SetValue之前
                Console.WriteLine("學生:" + delegation.Name + "\n");
                pdfService.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.Name, delegation.Name);

                Console.WriteLine("助手:" + delegation.Assistant + "\n");
                pdfService.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.Ass, delegation.Assistant);

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

        private void SetPdfFieldDelegationName(IDictionary<string, PdfFormField> fields, DelegationVM delegation,
            PdfDocument pdfDoc)
        {
            if (delegation.Date.Equals(lastDate) && delegation.Header.Contains(lastDelagation))
            {
                sameDelegationCount++;
                Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Header.Replace(lastDelagation, lastDelagation + sameDelegationCount) + "\n");
                pdfService.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" +
                    delegation.Header.Replace(lastDelagation, lastDelagation + sameDelegationCount));
            }
            else
            {
                sameDelegationCount = 1;
                Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Header + "\n");
                pdfService.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" +
                    delegation.Header);
            }
            lastDate = delegation.Date;
        }

        private void SetPdfFieldDelegationCell(IDictionary<string, PdfFormField> fields, DelegationVM delegation, 
            PdfDocument pdfDoc)
        {
            Console.WriteLine("委派:" + delegation.Header + "\n");
            if (delegation.Header.Contains(DelegationTypeName.Reading))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Reading);
                lastDelagation = DelegationTypeName.Reading;
            }
            else if (delegation.Header.Contains(DelegationTypeName.InitialCall))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.InitialCall);
                lastDelagation = DelegationTypeName.InitialCall;
            }
            else if (delegation.Header.Contains(DelegationTypeName.FirstRV))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FirstRV);
                lastDelagation = DelegationTypeName.FirstRV;
            }
            else if (delegation.Header.Contains(DelegationTypeName.SecondRV))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.SecondRV);
                lastDelagation = DelegationTypeName.SecondRV;
            }
            else if (delegation.Header.Contains(DelegationTypeName.BibleStudy))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BibleStudy);
                lastDelagation = DelegationTypeName.BibleStudy;
            }
            else if (delegation.Header.Contains(DelegationTypeName.BibleStudy2))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BibleStudy);
                lastDelagation = DelegationTypeName.BibleStudy2;
            }
            else if (delegation.Header.Contains(DelegationTypeName.Talk))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Talk);
                lastDelagation = DelegationTypeName.Talk;
            }
            //續訪需避免與第二次續訪衝突
            else if (delegation.Header.Contains(DelegationTypeName.FirstRV2) &&
                        !delegation.Header.Contains(DelegationTypeName.SecondRV))
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FirstRV);
                lastDelagation = DelegationTypeName.FirstRV2;
            }
            else
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Other);
                pdfService.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.OtherText, delegation.Header.Split('(')[0]);
            }
        }

        private void SetPdfFieldClass(IDictionary<string, PdfFormField> fields, DelegationVM delegation, 
            PdfDocument pdfDoc)
        {
            Console.WriteLine("班別:" + delegation.Class + "\n");
            if(delegation.Class == "1")
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Class1);
            }
            else if(delegation.Class == "2")
            {
                pdfService.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Class2);
            }
            else
            {
                Console.WriteLine("班別填錯了吧\n");
            }
        }      
    }
}
