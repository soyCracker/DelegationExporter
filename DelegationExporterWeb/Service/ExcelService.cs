using DelegationExporterWeb.Models;
using Microsoft.AspNetCore.Http;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;


namespace DelegationExporterWeb.Service
{
    public class ExcelService
    {
        public List<DelegationModel> ReadDelegation(IFormFile xlsFile)
        {
            IWorkbook workbook = GetIWorkbook(xlsFile);
            for (int i=0;i<=2;i++)
            {

            }
            return null;
        }

        private IWorkbook GetIWorkbook(IFormFile xlsFile)
        {
            MemoryStream ms = new MemoryStream();
            xlsFile.CopyToAsync(ms).ConfigureAwait(false);
            IWorkbook workbook;
            try
            {
                workbook = new XSSFWorkbook(ms);
            }
            catch (Exception)
            {
                workbook = new HSSFWorkbook(ms);
            }
            return workbook;
        }

        private List<DelegationModel> GetDelegationList(IWorkbook workbook)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            List<DelegationModel> delegationList = new List<DelegationModel>();
            for (int i=1;i<=2;i++)
            {

            }
            return null;
        }

        private List<DelegationModel> ReadOneClass(ISheet sheet, IWorkbook workbook, List<DelegationModel> delegationList, int classInt)
        {

            return null;
        }
    }
}
