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
        static void Main(string[] args)
        {
            //UpdateService updater = new UpdateService();
            //await updater.IsNewVer();

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
            Console.WriteLine("Work Start!!\n");
            DelegationService delegationService = new DelegationService();
            delegationService.Start();
            Console.WriteLine("-------------------------------\n");
            Console.WriteLine("Work End!!\n");
        }

        public static void InitEnvironment()
        {
            if(!Config.RELEASE_MODE)
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
