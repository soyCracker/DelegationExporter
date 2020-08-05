using DelegationExporterEntity.Entities;
using DelegationExporterWeb.Lock;
using DelegationExporterWeb.Util;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DelegationExporterWeb.Service
{
    public class NecessaryFileService
    {
        private DelegationExporterDBContext context;

        public NecessaryFileService(DelegationExporterDBContext context)
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

        public byte[] GetNecessaryFileData(string fileName)
        {
            using (var tryLock = new CacheLock())
            {
                if (CacheUtil.GetCache(fileName) != null)
                {
                    return (byte[])CacheUtil.GetCache(fileName);
                }
                else
                {
                    byte[] otaFile = context.NecessaryFile.FirstOrDefault(x => x.FileName == fileName).Data;
                    CacheUtil.SetCache(fileName, otaFile);
                    return otaFile;
                }
            }
        }
    }
}
