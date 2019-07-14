using DelegationExporter.Base;
using DelegationExporter.Except;
using DelegationExporter.Interface;
using DelegationExporter.Services;
using System;
using System.IO;

namespace DelegationExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            InitEnvironment();

            PrintInfo();

            Work();

            End();
        }

        public static void PrintInfo()
        {
            Console.WriteLine("Version:" + Config.VERSION + "\n");
            Console.WriteLine("Release Mode:" + Config.RELEASE_MODE + "\n");
            Console.WriteLine("Work Mode:" + ExternalConfig.Get().WorkMode + "\n");
            Console.WriteLine("Update Web url:" + Config.GIT_URL + "\n");
        }

        public static void Work()
        {
            try
            {
                Console.WriteLine("Work Start!!\n");
                IDelegationService delegationService = SelectWorkMode();
                delegationService.DoWork();
                Console.WriteLine("-------------------------------\n");
                Console.WriteLine("Work End!!\n");
            }
            catch(IOException)
            {
                Console.WriteLine("可能excel,pdf檔案不存在，或是檔案正被使用中請關閉檔案再試一次\n");
            }
            catch(NoFontException)
            {
                Console.WriteLine("字型檔案不存在，請檢查Font資料夾或重新下載程式\n");
            }
            catch(Exception ex)
            {
                Console.WriteLine("錯誤，請複製以下資訊給開發者:" + ex + "\n");
            }
        }

        public static void InitEnvironment()
        {
            if(!Config.RELEASE_MODE)
            {
                Directory.SetCurrentDirectory(Directory.GetCurrentDirectory() + "../../../../");
                Console.WriteLine("work dir:"+Directory.GetCurrentDirectory() + "\n");
            }          
            
            
        }

        public static void End()
        {
            Console.WriteLine("-------------------------------\n");
            Console.WriteLine("按兩次Enter結束程式");
            Console.ReadLine();
        }

        public static IDelegationService SelectWorkMode()
        {
            if(ExternalConfig.Get().WorkMode.Equals("TwSongshan"))
            {
                return new TwSongshanService();
            }
            else if(ExternalConfig.Get().WorkMode.Equals("TwNantou"))
            {
                return new TwNantouService();
            }
            return new TwSongshanService();
        }
    }
}
