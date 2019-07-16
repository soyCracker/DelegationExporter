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
        private string tempDate = "";
        private string tempXlsxFilePath = Config.FILE_FOLDER + "//" + Config.TEMP_NAME + ExternalConfig.Get().TargetXlsx;

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
                for (int i = 4; i <= sheet.LastRowNum; i++)
                {
                    SetXlsxList(sheet, i, s89List, 1);
                    SetXlsxList(sheet, i, s89List, 2);
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

        public void WriteDelegation<T>(T delegationT, string destFolder)
        {
            Console.WriteLine("-------------------------------\n");
            S89Xlsx delegation = delegationT as S89Xlsx;
            using (FileStream fs = new FileStream(Config.FILE_FOLDER + "//" + Config.PDF_FILE, FileMode.Open))
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fs), new PdfWriter(destFolder + "//" + delegation.Name + ".pdf"));
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                SetPdfField(fields, delegation);
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
                    s89Model.Date = tempDate;
                }
                else
                {
                    tempDate = s89Model.Date;
                }             
            }
            catch(Exception)
            {
                s89Model.Date = tempDate;
            }
            s89List.Add(s89Model);
        }

        private void SetPdfFieldDelegation(IDictionary<string, PdfFormField> fields, S89Xlsx delegation)
        {
            Console.WriteLine("委派:" + delegation.Delegation + "\n");
            if(delegation.Delegation.Contains("經文朗讀"))
            {
                fields[S89PdfField.Reading].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
            }
            else if(delegation.Delegation.Contains("初次交談"))
            {
                fields[S89PdfField.InitialCall].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
            }
            else if (delegation.Delegation.Contains("第一次續訪"))
            {
                fields[S89PdfField.FirstRV].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
            }
            else if (delegation.Delegation.Contains("第二次續訪"))
            {
                fields[S89PdfField.SecondRV].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
            }
            else if (delegation.Delegation.Contains("聖經研究") || delegation.Delegation.Contains("聖經討論"))
            {
                fields[S89PdfField.BibleStudy].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
            }
            else if (delegation.Delegation.Contains("演講"))
            {
                fields[S89PdfField.Talk].SetCheckType(PdfFormField.TYPE_CHECK).SetValue("true");
            }
        }

        private void SetPdfFieldClass(IDictionary<string, PdfFormField> fields, S89Xlsx delegation)
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

        private void SetPdfField(IDictionary<string, PdfFormField> fields, S89Xlsx delegation)
        {
            if (FontUtil.IsMsjhbdExist())
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

        public string BeforePrepareAndGetTempXlsx()
        {
            FileUtil.CopyTemp(Config.FILE_FOLDER + "//" + ExternalConfig.Get().TargetXlsx, tempXlsxFilePath);
            return tempXlsxFilePath;
        }
    }
}
