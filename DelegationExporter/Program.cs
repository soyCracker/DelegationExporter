using DelegationExporter.Base;
using DelegationExporter.Except;
using DelegationExporter.Services;
using System;
using System.IO;

namespace DelegationExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintInfo();

            Work();

            End();
        }

        public static void PrintInfo()
        {
            Console.WriteLine("Version:" + Config.VERSION + "\n");
            Console.WriteLine("Update Web url:" + Config.GIT_URL + "\n");
        }

        public static void Work()
        {
            try
            {
                Console.WriteLine("Work Start!!\n");
                S89chService s89 = new S89chService();
                s89.Work();
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

        public static void End()
        {
            Console.WriteLine("-------------------------------\n");
            Console.WriteLine("按兩次Enter結束程式");
            Console.ReadLine();
        }
    }
}
