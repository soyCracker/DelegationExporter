using Delegation.Service.Models;
using Delegation.Service.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Delegation.Service.Services
{
    public class RecordService
    {
        private string blockDate = "";

        public RecordService()
        {

        }

        public void Start(string outputFolder, string formFile, string assignFile)
        {
            string destFolder = SetAndGetFolder(outputFolder);

            List<DelegationVM> vmList = ReadDelegation(formFile);

            foreach (DelegationVM vm in vmList)
            {
                Console.WriteLine("test " + vm.Name);
            }

            Record(vmList, assignFile, destFolder);
        }

        private string SetAndGetFolder(string outputFolder)
        {
            if (!Directory.Exists(outputFolder + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(outputFolder + "//" + TimeUtil.GetTimeNow());
            }
            return outputFolder + "//" + TimeUtil.GetTimeNow();
        }

        public List<DelegationVM> ReadDelegation(string filePath)
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
                        SetDelegationList(sheet, i, vmList, j);
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

        private void SetDelegationList(ISheet sheet, int rowNum, List<DelegationVM> vmList, int classInt)
        {
            DelegationVM vm = new DelegationVM();
            try
            {
                vm.Name = sheet.GetRow(rowNum).GetCell(2 + (classInt - 1) * 2).ToString().Trim();
                if (string.IsNullOrEmpty(vm.Name))
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
                vm.Assistant = sheet.GetRow(rowNum).GetCell(3 + (classInt - 1) * 2).ToString();
            }
            catch (Exception)
            {
                vm.Assistant = "";
            }
            vm.Header = StringUtil.GetChinesePrintAble(sheet.GetRow(rowNum).GetCell(1).ToString());
            vm.Class = classInt.ToString();
            try
            {
                vm.Date = sheet.GetRow(rowNum).GetCell(0).ToString();
                if (string.IsNullOrEmpty(vm.Date))
                {
                    vm.Date = blockDate;
                }
                else
                {
                    blockDate = vm.Date;
                }
            }
            catch (Exception)
            {
                vm.Date = blockDate;
            }
            vmList.Add(vm);
        }

        private void Record(List<DelegationVM> vmList, string assignFilePath, string destFolder)
        {
            try
            {
                FileStream fs = new FileStream(assignFilePath, FileMode.Open);
                IWorkbook workbook;
                try
                {
                    workbook = new XSSFWorkbook(fs);
                }
                catch (Exception)
                {
                    workbook = new HSSFWorkbook(fs);
                }
                ISheet broSheet = workbook.GetSheetAt(0);
                ISheet sisSheet = workbook.GetSheetAt(1);
                ISheet assSheet = workbook.GetSheetAt(2);
                foreach (DelegationVM vm in vmList)
                {
                    if(!RecordBroSheet(broSheet, vm))
                    {
                        RecordSisSheet(sisSheet, assSheet, vm);
                    }
                }
                FileStream saveFs = new FileStream(Path.Combine(destFolder, "assignmentNext.xlsx"), FileMode.Create);
                workbook.Write(saveFs);
                workbook.Close();
                fs.Close();
                saveFs.Close();
            }
            catch (IOException)
            {
                throw;
            }
        }

        private bool RecordBroSheet(ISheet broSheet, DelegationVM vm)
        {
            int typeCol = 99;
            if (vm.Header.Contains(DelegationTypeName.Reading))
            {
                typeCol = 0;
            }
            else if (vm.Header.Contains(DelegationTypeName.Talk))
            {
                typeCol = 2;
            }
            else
            {
                return false;
            }

            for (int i = 2; i <= broSheet.LastRowNum; i++)
            {
                var row = broSheet.GetRow(i);
                if (row.GetCell(typeCol)!=null && row.GetCell(typeCol).ToString().Trim() == vm.Name)
                {
                    //順便整理一下string
                    row.GetCell(typeCol).SetCellValue(vm.Name);
                    if (row.GetCell(typeCol+1)==null)
                    {
                        row.CreateCell(typeCol+1);
                    }
                    row.GetCell(typeCol+1).SetCellValue(TimeUtil.CovertDateToFileNameStr(vm.Date));
                    return true;
                }
            }
            return false;
        }

        private bool RecordSisSheet(ISheet sisSheet, ISheet assSheet, DelegationVM vm)
        {
            for (int i = 1; i <= sisSheet.LastRowNum; i++)
            {
                var row = sisSheet.GetRow(i);
                if (row.GetCell(0)!=null && row.GetCell(0).ToString().Trim() == vm.Name)
                {
                    //順便整理一下string
                    row.GetCell(0).SetCellValue(vm.Name);
                    if (row.GetCell(1)==null)
                    {
                        row.CreateCell(1);
                    }
                    row.GetCell(1).SetCellValue(TimeUtil.CovertDateToFileNameStr(vm.Date));
                    string type = "X";
                    if(vm.Header.Contains(DelegationTypeName.InitialCall))
                    {
                        type = "初";
                    }
                    else if(vm.Header.Contains(DelegationTypeName.FirstRV) || vm.Header.Contains(DelegationTypeName.FirstRV2) || vm.Header.Contains(DelegationTypeName.SecondRV))
                    {
                        type = "續";
                    }
                    else if(vm.Header.Contains(DelegationTypeName.BibleStudy) || vm.Header.Contains(DelegationTypeName.BibleStudy2))
                    {
                        type = "課";
                    }
                    if (row.GetCell(2)==null)
                    {
                        row.CreateCell(2);
                    }
                    row.GetCell(2).SetCellValue(type);
                    RecordAssSheet(assSheet, vm);
                    return true;
                }
            }
            return false;
        }

        private void RecordAssSheet(ISheet assSheet, DelegationVM vm)
        {
            for (int i = 1; i <= assSheet.LastRowNum; i++)
            {
                var row = assSheet.GetRow(i);
                if (row.GetCell(0)!=null && row.GetCell(0).ToString().Trim() == vm.Assistant)
                {
                    //順便整理一下string
                    row.GetCell(0).SetCellValue(vm.Assistant);
                    if (row.GetCell(1)==null)
                    {
                        row.CreateCell(1);
                    }
                    row.GetCell(1).SetCellValue(TimeUtil.CovertDateToFileNameStr(vm.Date));
                    for(int j=2;j<=7;j++)
                    {
                        if(row.GetCell(j)==null)
                        {
                            row.CreateCell(j);
                            row.GetCell(j).SetCellValue(vm.Name);
                            return;
                        }
                        if(string.IsNullOrEmpty(row.GetCell(j).ToString().Trim()))
                        {
                            row.GetCell(j).SetCellValue(vm.Name);
                            return;
                        }
                        if(j==7)
                        {
                            row.GetCell(2).SetCellValue(vm.Name);
                            for(int k=3;k<=7;k++)
                            {
                                row.GetCell(k).SetCellValue("");
                            }
                            return;
                        }
                    }
                }
            }
        }
    }
}
