using DelegationExporter.Base;
using DelegationExporter.Services;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DelegationExporter
{
    class Program
    {
        static async Task Main(string[] args)
        {
            await PrintInfo();

            InitEnvironment();

            Work();

            End();
        }

        public async static Task PrintInfo()
        {
            UpdateService updater = new UpdateService();
            Console.WriteLine("The Latest Version: " + await updater.GetLatestVerInfo(Constant.GIT_RELEASE_API) + "\n");
            Console.WriteLine("Version:" + Constant.VERSION + "\n");
            Console.WriteLine("Release Mode:" + Constant.RELEASE_MODE + "\n");
        }

        public static void Work()
        {
            Console.WriteLine("Work Start!!\n");
            DelegationService delegationService = new DelegationService();
            delegationService.Start();
            Console.WriteLine("-------------------------------\n");
            Console.WriteLine("Work Finish!\n");
        }

        public static void InitEnvironment()
        {
            if(!Constant.RELEASE_MODE)
            {
                Directory.SetCurrentDirectory(Directory.GetCurrentDirectory() + "../../../../");
                Console.WriteLine("work dir:"+Directory.GetCurrentDirectory() + "\n");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.Unicode;
        }

        public static void End()
        {
            Console.WriteLine("-------------------------------\n");
            Console.WriteLine("按兩次Enter結束程式");
            Console.ReadLine();
        }
    }
}
