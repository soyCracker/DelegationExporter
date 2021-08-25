using System;
using System.Collections.Generic;
using System.Text;

namespace DelegationExporter.Util
{
    public class TimeUtil
    {
        public static string GetTimeNow()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        public static string CovertDateToFileNameStr(string dateStr)
        {
            return DateTime.Parse(dateStr).ToString("yyyyMMdd");
        }
    }
}
