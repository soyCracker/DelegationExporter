using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Threading.Tasks;

namespace DelegationExporterWeb.Util
{
    public class CacheUtil
    {
        private static ObjectCache cache = MemoryCache.Default;
        private readonly static int expiration = 5; //minute
        private readonly static int cacheSizeLimit = 50; //MB

        //cache strategy : if cache size over limit, do not put data in cache, and waiting for cache expired naturally
        public static void SetCache(string key, object obj)
        {
            if (!IsCacheOverLimit())
            {
                var policy = new CacheItemPolicy();
                // 過期時間內未使用快取時，回收快取
                policy.SlidingExpiration = TimeSpan.FromMinutes(expiration);
                cache.Set(key, obj, policy);
            }
        }

        public static object GetCache(string key)
        {
            return cache[key];
        }

        public static double GetCacheMBSize()
        {
            double size = 0;
            foreach (var item in cache)
            {
                size += ((byte[])item.Value).Length;
            }
            return size / (1024 * 1024);
        }

        public static bool IsCacheOverLimit()
        {
            if (GetCacheMBSize() >= cacheSizeLimit)
            {
                return true;
            }
            return false;
        }
    }
}
