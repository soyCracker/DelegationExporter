using Delegation.Service.Except;
using DelegationConsoleTool.Base;
using System.Text;

namespace DelegationConsoleTool.Kits
{
    public class EnvirKit
    {
        public void InitEnvironment()
        {
            if (!Constant.RELEASE_MODE)
            {
                Directory.SetCurrentDirectory(Directory.GetCurrentDirectory() + "../../../../");
                Console.WriteLine("work dir:" + Directory.GetCurrentDirectory() + "\n");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.Unicode;
        }

        public string PrepareAndGetTempXlsx(string fileFolder)
        {
            string[] files = Directory.GetFiles(fileFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).ToArray();
            if (files.Length > 1)
            {
                throw new MultipleXlsException();
            }
            else if (files.Length == 1)
            {
                string file = files[0];
                string tempXls = file + "TEMPX";
                if (File.Exists(tempXls))
                {
                    File.Delete(tempXls);
                }
                File.Copy(file, tempXls);
                Console.WriteLine("PrepareAndGetTempXlsx() tempXls:" + tempXls + "\n");
                return tempXls;

            }
            return "";
        }

        public string GetAssignRecordXlsx(string fileFolder)
        {
            string[] files = Directory.GetFiles(fileFolder).Where(f => f.EndsWith(".xls") || f.EndsWith(".xlsx")).ToArray();
            string file = files[0];
            return file;
        }
    }
}
