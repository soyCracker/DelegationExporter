using DelegationExporter.Base;
using DelegationExporter.Except;
using DelegationExporter.Interface;
using DelegationExporter.Model;
using DelegationExporter.Util;
using System;
using System.Collections.Generic;
using System.IO;

namespace DelegationExporter.Services
{
    public class DelegationService
    {
        private IDelegationService delegationService;

        public DelegationService()
        {
            delegationService = SelectWorkMode();
        }

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
                Console.WriteLine("Update Web url:" + Config.GIT_URL + "\n");
                Console.WriteLine(ex + "\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine("未知錯誤，請複製以下資訊給開發者:" + ex + "\n");
            }
            
        }

        private void Process()
        {
            string destFolder = GetFolder();

            string tempXlsxFilePath = delegationService.BeforePrepareAndGetTempXlsx();

            List<S89Xlsx> list = delegationService.ReadDelegation<S89Xlsx>(tempXlsxFilePath);

            delegationService.WriteDelegation(list, destFolder);

            AfterClear(tempXlsxFilePath);
        }

        private IDelegationService SelectWorkMode()
        {
            if (ExternalConfig.Get().WorkMode.Equals("TwSongshan"))
            {
                return new TwSongshanService();
            }
            else if (ExternalConfig.Get().WorkMode.Equals("TwNantou"))
            {
                return new TwNantouService();
            }
            return new TwSongshanService();
        }

        private string GetFolder()
        {
            if (!Directory.Exists(Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow());
            }
            return Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow();
        }
        private void AfterClear(string tempXlsxFilePath)
        {
            FileUtil.ClearTemp(tempXlsxFilePath);
        }
    }
}
