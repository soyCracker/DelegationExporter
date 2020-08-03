using DelegationExporterEntity.Entities;
using DelegationExporterWeb.Base;
using DelegationExporterWeb.Models;
using DelegationExporterWeb.Util;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DelegationExporterWeb.Service
{
    public class PdfServce
    {
        private DelegationExporterDBContext context;
        private int sameSubjectCount = 0;
        private string previousDate = "";
        private string previousSubject = "";

        public PdfServce(DelegationExporterDBContext context)
        {
            this.context = context;
        }

        public void SaveFileToDB(IFormFile file)
        {       
            using (var ms = new MemoryStream())
            {
                file.CopyTo(ms);
                byte[] byteFile = ms.ToArray();
                NecessaryFile necessaryFile = new NecessaryFile();
                necessaryFile.FileName = file.FileName;
                necessaryFile.Data = byteFile;
                necessaryFile.UpdateTime = DateTime.UtcNow;
                context.NecessaryFile.Add(necessaryFile);
                context.SaveChanges();
            }            
        }

        public void WriteInDelegation(List<DelegationModel> delegationList)
        {
            foreach(DelegationModel delegationModel in delegationList)
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(WhichS89(delegationModel)), new PdfWriter(GetPdfFileName(delegationModel.Name)));
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                SetPdfField(fields, delegationModel, pdfDoc);
                pdfDoc.Close();
            }
        }

        private MemoryStream WhichS89(DelegationModel delegationModel)
        {
            NecessaryFile necessaryFile;
            if (delegationModel.Name.Contains(Constant.J_STR))
            {
                necessaryFile = context.NecessaryFile.FirstOrDefault(x => x.FileName == Constant.S89J_FILE_NAME);
            }
            else
            {
                necessaryFile = context.NecessaryFile.FirstOrDefault(x => x.FileName == Constant.S89CH_FILE_NAME);
            }
            return new MemoryStream(necessaryFile.Data);           
        }

        private string GetPdfFileName(string name)
        {
            string temp = Constant.DEFAULT_EXPORT_PDF_FILE_NAME;
            return temp.Replace(Constant.LANGUAGE_STR, name.Contains(Constant.J_STR) ? "-日語" : "")
                .Replace(Constant.NAME_STR, name.Replace(Constant.J_STR, "")).Replace(Constant.DATE_STR, TimeUtil.GetTimeNow());
        }

        private void SetPdfField(IDictionary<string, PdfFormField> fields, DelegationModel delegation, PdfDocument pdfDoc)
        {
            //我她X的，SetFont要在SetValue之前
            PdfUtil.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.NAME, delegation.Name, context);

            PdfUtil.SetPdfFeldValueCenter(fields, pdfDoc, S89PdfField.ASSISTANT, delegation.Assistant, context);

            SetPdfFieldTitle(fields, delegation, pdfDoc);

            SetPdfFieldSubjectCell(fields, delegation, pdfDoc);

            SetPdfFieldClass(fields, delegation, pdfDoc);
        }

        private void SetPdfFieldTitle(IDictionary<string, PdfFormField> fields, DelegationModel delegation, PdfDocument pdfDoc)
        {
            if (delegation.Date.Equals(previousDate) && delegation.Subject.Contains(previousSubject))
            {
                sameSubjectCount++;
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.DATE, delegation.Date + "-" + 
                    delegation.Subject.Replace(delegation.Subject, delegation.Subject + sameSubjectCount), context);
            }
            else
            {
                sameSubjectCount = 1;
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.DATE, delegation.Date + "-" + delegation.Subject, context);
            }
            previousDate = delegation.Date;
        }

        private void SetPdfFieldSubjectCell(IDictionary<string, PdfFormField> fields, DelegationModel delegation, PdfDocument pdfDoc)
        {
            if (delegation.Subject.Contains(DelegationType.READING))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.READING, context);
                previousSubject = DelegationType.READING;
            }
            else if (delegation.Subject.Contains(DelegationType.INITIAL_CALL))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.INITIAL_CALL, context);
                previousSubject = DelegationType.INITIAL_CALL;
            }
            else if (delegation.Subject.Contains(DelegationType.FIRST_RV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FIRST_RV, context);
                previousSubject = DelegationType.FIRST_RV;
            }
            else if (delegation.Subject.Contains(DelegationType.SECOND_RV))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.SECOND_RV, context);
                previousSubject = DelegationType.SECOND_RV;
            }
            else if (delegation.Subject.Contains(DelegationType.BIBLE_STUDY))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BIBLE_STUDY, context);
                previousSubject = DelegationType.BIBLE_STUDY;
            }
            else if (delegation.Subject.Contains(DelegationType.BIBLE_STUDY2))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.BIBLE_STUDY, context);
                previousSubject = DelegationType.BIBLE_STUDY2;
            }
            else if (delegation.Subject.Contains(DelegationType.TALK))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.TALK, context);
                previousSubject = DelegationType.TALK;
            }
            //續訪擺在最後避免與第二次續訪衝突
            else if (delegation.Subject.Contains(DelegationType.FIRST_RV2))
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.FIRST_RV, context);
                previousSubject = DelegationType.FIRST_RV2;
            }
            else
            {
                PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.OTHER, context);
                PdfUtil.SetPdfFeldValueSmall(fields, pdfDoc, S89PdfField.OTHER_TEXT, delegation.Subject.Split('(')[0], context);
            }
        }

        private void SetPdfFieldClass(IDictionary<string, PdfFormField> fields, DelegationModel delegation, PdfDocument pdfDoc)
        {
            switch (delegation.DelegationClass)
            {
                case "1":
                    PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.CLASS1, context);
                    break;
                case "2":
                    PdfUtil.SetPdfCheckBoxSelected(fields, pdfDoc, S89PdfField.CLASS2, context);
                    break;
                default:
                    Console.WriteLine("班別填錯了吧\n");
                    break;
            }
        }
    }
}
