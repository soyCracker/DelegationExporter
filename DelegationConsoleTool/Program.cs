using Delegation.Service.Except;
using Delegation.Service.Services;
using DelegationConsoleTool.Base;
using System;
using System.IO;
using System.Text;

namespace DelegationConsoleTool
{
    class Program
    {
        static string delegationFormFolder = Constant.FILE_FOLDER + @"\" + Constant.DELEGATION_FORM_FOLDER;
        static string assignmentFolder = Constant.FILE_FOLDER + @"\" + Constant.ASSIGNMENT_FOLDER;

        static void Main(string[] args)
        {
            InitEnvironment();                     
            while(true)
            {
                FucDescOutput();
                string fucKey = Console.ReadLine();
                Fuchub(fucKey);
            }          
        }

        static void InitEnvironment()
        {
            //Debug模式，執行環境切換到專案資料夾
            if (!Constant.RELEASE_MODE)
            {
                Directory.SetCurrentDirectory(Directory.GetCurrentDirectory() + "../../../../");
                Console.WriteLine("work dir:" + Directory.GetCurrentDirectory() + "\n");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.Unicode;
        }

        static void FucDescOutput()
        {
            Console.WriteLine("Delegation Console Tool 功能:");
            Console.WriteLine("委派紀錄填寫 - 1");
            Console.WriteLine("輸出委派單 - 3");
            Console.Write("輸入功能:");
        }

        static void Fuchub(string fucKey)
        {
            if (fucKey == "1")
            {
                RecordWork();
            }
            else if (fucKey == "3")
            {
                ExportWork();
            }
            else if(fucKey=="99")
            {
                TestWork();
            }
        }

        static void RecordWork()
        {
            Console.WriteLine("委派紀錄填寫");
            RecordService recordService = new RecordService();
            recordService.Start(Constant.OUTPUT_FOLDER, delegationFormFolder, Constant.TEMP_NAME, assignmentFolder);
        }

        static void ExportWork()
        {
            try
            {
                Console.WriteLine("輸出委派單");
                ExportService exportService = new ExportService(Constant.FONT_FOLDER);
                exportService.Start(Constant.OUTPUT_FOLDER, delegationFormFolder, Constant.TEMP_NAME, Constant.S89CH, 
                    Constant.S89J, Constant.DESC_STR, Constant.DESC_JP_STR, Constant.JP_FLAG_STR);
            }
            catch (IOException ex)
            {
                Console.WriteLine("可能excel或pdf檔案不存在、config內的檔名填錯，或是檔案正被使用中請關閉所有檔案再試一次\n");
                Console.WriteLine(ex + "\n");
            }
            catch (NoFontException ex)
            {
                Console.WriteLine("字型檔案不存在，請檢查Font資料夾或重新下載程式\n");
                Console.WriteLine("Update Web url:" + Constant.GIT_RELEASE_API + "\n");
                Console.WriteLine(ex + "\n");
            }
            catch (MultipleXlsException ex)
            {
                Console.WriteLine("存在多個xls或xlsx於資料夾中\n");
                Console.WriteLine(ex + "\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine("未知錯誤，請複製以下資訊給開發者:" + ex + "\n");
            }
        }

        static void TestWork()
        {
            Console.WriteLine("TestWork() !!!\n");
        }
    }
}






