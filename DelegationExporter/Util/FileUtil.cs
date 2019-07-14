using DelegationExporter.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DelegationExporter.Util
{
    public class FileUtil
    {
        public static List<string> GetFileList(string dir)
        {
            List<string> files = new List<string>();
            try
            {              
                foreach (string f in Directory.GetFiles(dir))
                {
                    Console.WriteLine(f);
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
