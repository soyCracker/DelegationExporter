using DelegationExporter.Base;
using DelegationExporter.Except;
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
using System.Linq;

namespace DelegationExporter.Services
{
    public class DelegationService
    {
        private string blockDate = "";
        private string lastDelagation = "";
        private string lastDate = "";
        private int sameDelegationCount = 1;

        public void Start()
        {
            try
            {
                Process();
            }
            catch (IOException ex)
            {
                Console.WriteLine("可能excel或pdf檔案不存在、config內的檔名填錯，或是檔案正被使用中請關閉所有檔案再試一次\n");
                Console.WriteLine(ex + "\n");
            }
            catch (NoFontException ex)
            {
                Console.WriteLine("字型檔案不存在，請檢查Font資料夾或重新下載程式\n");
                Console.WriteLine("Update Web url:" + Constant.GIT_URL + "\n");
                Console.WriteLine(ex + "\n");
            }
            catch(MultipleXlsException ex)
            {
                Console.WriteLine("存在多個xls或xlsx於資料夾中\n");
                Console.WriteLine(ex + "\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine("未知錯誤，請複製以下資訊給開發者:" + ex + "\n");
            }
            
        }

        private void Process()
        {
            string destFolder = SetAndGetFolder();

            string tempXlsxFilePath = PrepareAndGetTempXlsx();

            List<S89Xlsx> list = ReadDelegation<S89Xlsx>(tempXlsxFilePath);

            WriteDelegation(list, destFolder);

            FileUtil.ClearTemp(tempXlsxFilePath);
        }

        private string SetAndGetFolder()
        {
            if (!Directory.Exists(Constant.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(Constant.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow());
            }
            return Constant.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow();
        }  

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
                for (int j = 1; j <= 2; j++)
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
                string s89 = Constant.S89CH;
                string description = Constant.DESCRIP_STR;
                if (delegation.Name.Contains(Constant.DISTINT_JP_STR))
                {
                    s89 = Constant.S89J;
                    description = description + Constant.DESCRIP_JP_STR;
                    delegation.Name = delegation.Name.Replace(Constant.DISTINT_JP_STR, "");
                }
                WriteInPdf(s89, delegation, destFolder, description);
            }
        }

        private void WriteInPdf(string s89, S89Xlsx delegation, string destFolder, string description)
        {
            using (FileStream fs = new FileStream(Constant.FILE_FOLDER + "//" + s89, FileMode.Open))
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(PdfUtil.FileNameExistAddR(destFolder + "//" + description + delegation.Name + Constant.PDF_FILE_NAME_EXTENSION)));
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
                if (string.IsNullOrEmpty(s89Model.Name))
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
                s89Model.Assistant = sheet.GetRow(rowNum).GetCell(3 + (classInt - 1) * 2).ToString();
            }
            catch (Exception)
            {
                s89Model.Assistant = "";
            }
            s89Model.Header = StringUtil.GetChinesePrintAble(sheet.GetRow(rowNum).GetCell(1).ToString());
            s89Model.Class = classInt.ToString();
            try
            {
                s89Model.Date = sheet.GetRow(rowNum).GetCell(0).ToString();
                if (string.IsNullOrEmpty(s89Model.Date))
                {
                    s89Model.Date = blockDate;
                }
                else
                {
                    blockDate = s89Model.Date;
                }
            }
            catch (Exception)
            {
                s89Model.Date = blockDate;
            }
            s89List.Add(s89Model);
        }

        private void SetPdfFieldDelegationCell(IDictionary<string, PdfFormField> fields, S89Xlsx delegation, PdfDocument pdfDoc)
        {
            Console.WriteLine("委派:" + delegation.Header + "\n");
            if (delegation.Header.Contains(S89chDelegationName.Reading))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Reading);
                lastDelagation = S89chDelegationName.Reading;
            }
            else if (delegation.Header.Contains(S89chDelegationName.InitialCall))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.InitialCall);
                lastDelagation = S89chDelegationName.InitialCall;
            }
            else if (delegation.Header.Contains(S89chDelegationName.FirstRV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FirstRV);
                lastDelagation = S89chDelegationName.FirstRV;
            }
            else if (delegation.Header.Contains(S89chDelegationName.SecondRV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.SecondRV);
                lastDelagation = S89chDelegationName.SecondRV;
            }
            else if (delegation.Header.Contains(S89chDelegationName.BibleStudy))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BibleStudy);
                lastDelagation = S89chDelegationName.BibleStudy;
            }
            else if (delegation.Header.Contains(S89chDelegationName.BibleStudy2))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BibleStudy);
                lastDelagation = S89chDelegationName.BibleStudy2;
            }
            else if (delegation.Header.Contains(S89chDelegationName.Talk))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Talk);
                lastDelagation = S89chDelegationName.Talk;
            }
            //續訪需避免與第二次續訪衝突
            else if (delegation.Header.Contains(S89chDelegationName.FirstRV2) &&
                        !delegation.Header.Contains(S89chDelegationName.SecondRV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FirstRV);
                lastDelagation = S89chDelegationName.FirstRV2;
            }
            else
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.Other);
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.OtherText, delegation.Header.Split('(')[0]);
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
            if (delegation.Date.Equals(lastDate) && delegation.Header.Contains(lastDelagation))
            {
                sameDelegationCount++;
                Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Header.Replace(lastDelagation, lastDelagation + sameDelegationCount) + "\n");
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" + delegation.Header.Replace(lastDelagation, lastDelagation + sameDelegationCount));
            }
            else
            {
                sameDelegationCount = 1;
                Console.WriteLine("日期:" + delegation.Date + "-" + delegation.Header + "\n");
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.Date, delegation.Date + "-" + delegation.Header);
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

        public string PrepareAndGetTempXlsx()
        {
            List<string> fileList = FileUtil.GetFileList(Constant.FILE_FOLDER).FindAll(x => x.ToLowerInvariant().EndsWith(".xls") || x.ToLowerInvariant().EndsWith(".xlsx"));
            if(fileList.Count() > 1)
            {
                throw new MultipleXlsException();
            }
            string file = fileList.First();
            string tempXls = file + Constant.TEMP_NAME;
            if (File.Exists(tempXls))
            {
                File.Delete(tempXls);
            }
            FileUtil.CopyTemp(file, tempXls);           
            Console.WriteLine("PrepareAndGetTempXlsx() tempXls:" + tempXls + "\n");
            return tempXls;
        }
    }
}
