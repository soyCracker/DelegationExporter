using Delegation.Service.Except;
using Delegation.Service.Utils;
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
                string fileExt = file.ToLower().EndsWith("xlsx") ? "xlsx" : "xls";
                string tempXls = string.Format("{0}_temp.{1}", file, fileExt);
                if (File.Exists(tempXls))
                {
                    File.Delete(tempXls);
                }
                File.Copy(file, tempXls);
                Console.WriteLine("tempXls:" + tempXls + "\n");
                return tempXls;
            }
            return "";
        }

        public string CreateOutputFolder(string outputFolder)
        {
            string[] pathParam = { outputFolder, TimeUtil.GetTimeNow() };
            string targetPath = Path.Combine(pathParam);
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }
            return targetPath;
        }
    }
}
