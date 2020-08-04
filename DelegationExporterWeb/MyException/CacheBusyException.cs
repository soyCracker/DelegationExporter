using System;

namespace DelegationExporterWeb.MyException
{
    public class CacheBusyException : Exception
    {
        public CacheBusyException() : base("CacheBusyException")
        {

        }
    }
}
