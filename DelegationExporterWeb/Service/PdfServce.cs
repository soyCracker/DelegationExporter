using DelegationExporterEntity.Entities;
using DelegationExporterWeb.Base;
using DelegationExporterWeb.Models;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DelegationExporterWeb.Service
{
    public class PdfServce
    {
        private DelegationExporterDBContext context;

        public PdfServce(DelegationExporterDBContext context)
        {
            this.context = context;
        }

        public void SavePdfToDB(IFormFile pdf)
        {       
            using (var ms = new MemoryStream())
            {
                pdf.CopyTo(ms);
                byte[] byteFile = ms.ToArray();
                NecessaryFile necessaryFile = new NecessaryFile();
                necessaryFile.FileName = pdf.FileName;
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
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(WhichS89(delegationModel)), new PdfWriter(""));
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                IDictionary<string, PdfFormField> fields = form.GetFormFields();
                //SetPdfField(fields, delegation, pdfDoc);
                pdfDoc.Close();
            }
        }

        public MemoryStream WhichS89(DelegationModel delegationModel)
        {
            NecessaryFile necessaryFile;
            if (delegationModel.Name.Contains("J"))
            {
                necessaryFile = context.NecessaryFile.FirstOrDefault(x => x.FileName == Constant.S89J_FILE_NAME);
            }
            else
            {
                necessaryFile = context.NecessaryFile.FirstOrDefault(x => x.FileName == Constant.S89CH_FILE_NAME);
            }
            return new MemoryStream(necessaryFile.Data);           
        }
    }
}
