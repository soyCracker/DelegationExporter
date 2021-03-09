using DelegationExporter.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DelegationExporter.Util
{
    public class FileUtil
    {
        public static void CopyTemp(string sourcePath, string newTempPath)
        {
            try
            {
                File.Copy(sourcePath, newTempPath);
            }
            catch (IOException)
            {
                throw;
            }
        }

        public static void ClearTemp(string tempPath)
        {
            try
            {
                File.Delete(tempPath);
            }
            catch (IOException)
            {
                throw;
            }
        }

        public static string GetOutputFolder()
        {
            if (!Directory.Exists(Constant.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(Constant.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow());
            }
            return Constant.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow();
        }

        public static List<string> GetFileList(string dir)
        {
            List<string> files = new List<string>();
            try
            {              
                foreach (string f in Directory.GetFiles(dir))
                {
                    //Console.WriteLine("FileUtil GetFileList() file:" + f);
                    files.Add(f);
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
            return files;
        }
    }
}
