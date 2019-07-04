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
    }
}
