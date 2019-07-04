using DelegationExporter.Base;
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
            catch (IOException)
            {
                Console.WriteLine("請關閉excel再試一次\n");
                throw;
            }
        }
    }
}
