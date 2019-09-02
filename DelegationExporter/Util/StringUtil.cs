using System;

namespace DelegationExporter.Util
{
    public class StringUtil
    {
        /// <summary>
        /// get chinese or printable string
        /// </summary>
        /// <param name="source"></param>
        /// <returns>chinese or printable string</returns>
        public static string GetChinesePrintAble(string source)
        {
            try
            {
                string result = "";
                for (int i = 0; i < source.Length; i++)
                {
                    if (char.ConvertToUtf32(source, i) >= Convert.ToInt32("4e00", 16) &&
                        char.ConvertToUtf32(source, i) <= Convert.ToInt32("9fff", 16) ||
                        (source[i] >= 32 && source[i] <= 126))
                    {
                        result += source.Substring(i, 1);
                    }
                }
                return result;
            }
            catch(Exception)
            {
                return source;
            }
        }

        public static bool IsContainEveryChar(string target, string keys)
        {
            try
            {
                foreach(char c in keys)
                {
                    if(!target.Contains(c))
                    {
                        //Console.WriteLine("false:"+c);
                        return false;
                    }
                }
                //Console.WriteLine("true");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }            
        }
    }
}
