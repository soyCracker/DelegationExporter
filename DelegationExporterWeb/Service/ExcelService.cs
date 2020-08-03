using DelegationExporterWeb.Models;
using DelegationExporterWeb.Util;
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
        private string blockDate = "";

        public List<DelegationModel> ReadDelegation(IFormFile xlsFile)
        {
            IWorkbook workbook = GetIWorkbook(xlsFile);
            return GetDelegationList(workbook);           
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
                delegationList.AddRange(ReadOneClass(sheet, i));
            }
            return delegationList;
        }

        private List<DelegationModel> ReadOneClass(ISheet sheet, int classInt)
        {
            List<DelegationModel> tempList = new List<DelegationModel>();
            for (int i = 4; i <= sheet.LastRowNum; i++)
            {
                DelegationModel delegationModel = new DelegationModel();
                try
                {
                    delegationModel.Name = sheet.GetRow(i).GetCell(2 + (classInt - 1) * 2).ToString();
                    if (string.IsNullOrEmpty(delegationModel.Name))
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                try
                {
                    delegationModel.Assistant = sheet.GetRow(i).GetCell(3 + (classInt - 1) * 2).ToString();
                }
                catch (Exception)
                {
                    delegationModel.Assistant = "";
                }
                delegationModel.Subject = StringUtil.GetChinesePrintAble(sheet.GetRow(i).GetCell(1).ToString());
                delegationModel.DelegationClass = classInt.ToString();
                try
                {
                    delegationModel.Date = sheet.GetRow(i).GetCell(0).ToString();
                    if (string.IsNullOrEmpty(delegationModel.Date))
                    {
                        delegationModel.Date = blockDate;
                    }
                    else
                    {
                        blockDate = delegationModel.Date;
                    }
                }
                catch (Exception)
                {
                    delegationModel.Date = blockDate;
                }
                tempList.Add(delegationModel);
            }
            return tempList;
        }
    }
}
