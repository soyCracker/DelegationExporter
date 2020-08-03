using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DelegationExporterWeb.Util
{
    public class TimeUtil
    {
        public static string GetTimeNow()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }
    }
}
