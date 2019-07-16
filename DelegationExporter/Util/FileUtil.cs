﻿using DelegationExporter.Base;
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
            if (!Directory.Exists(Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow()))
            {
                Directory.CreateDirectory(Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow());
            }
            return Config.OUTPUT_FOLDER + "//" + TimeUtil.GetTimeNow();
        }

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
