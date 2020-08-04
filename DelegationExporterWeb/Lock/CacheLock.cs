using DelegationExporterWeb.MyException;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace DelegationExporterWeb.Lock
{
    public class CacheLock : IDisposable
    {
        private static readonly object mutex = new object();

        public CacheLock()
        {
            if (Monitor.TryEnter(mutex) == false)
            {
                throw new CacheBusyException();
            }
        }

        public void Dispose()
        {
            Monitor.Exit(mutex);
        }
    }
}
