using DelegationExporterEntity.Entities;
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
    }
}
