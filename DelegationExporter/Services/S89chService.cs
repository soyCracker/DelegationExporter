using DelegationExporter.Base;
using DelegationExporter.Except;
using DelegationExporter.Interface;
using DelegationExporter.Model;
using DelegationExporter.Util;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace DelegationExporter.Services
{
    public class S89chService : IDelegationService<SongSanDelegationXlsxModel>
    {
        public void Work()
        {
            string destFolder = GetFolder();
            ReadProcess(destFolder, Config.BRO_SHEET);           
            ReadProcess(destFolder, Config.SIS_SHEET);
        }

        private void ReadProcess(string destFolder, string genderSheet)
        {
            List<SongSanDelegationXlsxModel> list = ReadDelegation(Config.FILE_FOLDER + "//" + Config.TARGET_XLSX, genderSheet);
            foreach (SongSanDelegationXlsxModel element in list)
            {
                WriteDelegation(element, destFolder);
            }
        }

        private string GetFolder()
        {
            if (!Directory.Exists(Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow());
            }
            return Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow();
        }

        private void SetXlsxList(ISheet sheet, int rowNum, List<SongSanDelegationXlsxModel> s89List)
        {
            if (sheet.GetRow(rowNum).GetCell(5).ToString().Equals("V") || sheet.GetRow(rowNum).GetCell(5).ToString().Equals("v"))
            {
                SongSanDelegationXlsxModel s89Model = new SongSanDelegationXlsxModel();
                s89Model.Name = sheet.GetRow(rowNum).GetCell(0).ToString();
                try
                {
                    s89Model.Assistant = sheet.GetRow(rowNum).GetCell(1).ToString();
                }
                catch (Exception)
                {
                    s89Model.Assistant = "";
                }
                s89Model.Delegation = sheet.GetRow(rowNum).GetCell(2).ToString();
                s89Model.Class = sheet.GetRow(rowNum).GetCell(3).ToString();
                s89Model.Date = sheet.GetRow(rowNum).GetCell(4).DateCellValue.ToString("yyyy/MM/dd");
                s89List.Add(s89Model);
            }
        }

        

        private void SetPdfFieldDelegation(IDictionary<string, PdfFormField> fields, SongSanDelegationXlsxModel delegation)
        {
            Console.WriteLine("委派:" + delegation.Delegation + "\n");
            switch (delegation.Delegation)
            {
                case "經文朗讀":
                    fields[S89PdfField.Reading].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                case "初次交談":
                    fields[S89PdfField.InitialCall].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                case "第一次續訪":
                    fields[S89PdfField.FirstRV].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                case "第二次續訪":
                    fields[S89PdfField.SecondRV].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                case "聖經研究":
                    fields[S89PdfField.BibleStudy].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                case "演講":
                    fields[S89PdfField.Talk].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                default:
                    Console.WriteLine("委派填錯了吧/n");
                    break;
            }
        }

        private void SetPdfFieldClass(IDictionary<string, PdfFormField> fields, SongSanDelegationXlsxModel delegation)
        {
            Console.WriteLine("班別:" + delegation.Class + "\n");
            switch (delegation.Class)
            {
                case "1":
                    fields[S89PdfField.Class1].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                case "2":
                    fields[S89PdfField.Class2].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
                    break;
                default:
                    Console.WriteLine("班別填錯了吧\n");
                    break;
            }
        }

        private void SetPdfField(IDictionary<string, PdfFormField> fields, SongSanDelegationXlsxModel delegation)
        {
            if(FontUtil.IsMsjhbdExist())
            {
                //我她X的，SetFont要在SetValue之前
                Console.WriteLine("學生:" + delegation.Name + "\n");
                fields[S89PdfField.Name].SetFont(FontUtil.GetMsjhbdFont()).SetValue(delegation.Name);
                Console.WriteLine("助手:" + delegation.Assistant + "\n");
                fields[S89PdfField.Ass].SetFont(FontUtil.GetMsjhbdFont()).SetValue(delegation.Assistant);
                Console.WriteLine("日期:" + delegation.Date + "\n");
                fields[S89PdfField.Date].SetFont(FontUtil.GetMsjhbdFont()).SetValue(delegation.Date);
                SetPdfFieldDelegation(fields, delegation);
                SetPdfFieldClass(fields, delegation);
            }
            else
            {
                throw new NoFontException();
            }
        }

        public List<SongSanDelegationXlsxModel> ReadDelegation(string filePath, string sheetName)
        {
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Open);
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheet(sheetName);
                List<SongSanDelegationXlsxModel> s89List = new List<SongSanDelegationXlsxModel>();
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        SetXlsxList(sheet, i, s89List);
                    }
                    catch (NullReferenceException)
                    {
                        //沒有要export
                    }
                }
                workbook.Close();
                fs.Close();
                return s89List;
            }
            catch (IOException)
            {
                throw;
            }
        }

        public void WriteDelegation(SongSanDelegationXlsxModel delegation, string destFolder)
        {
            Console.WriteLine("-------------------------------\n");
            using (FileStream fs = new FileStream(Config.FILE_FOLDER + "//" + Config.PDF_FILE, FileMode.Open))
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(destFolder + "//" + delegation.Name + ".pdf")); 
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                SetPdfField(fields, delegation);
                pdfDoc.Close();
            }
        }
    }
}
