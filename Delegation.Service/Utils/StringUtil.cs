using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Delegation.Service.Utils
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
            catch (Exception)
            {
                return source;
            }
        }
    }
}
