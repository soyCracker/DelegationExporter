using DelegationExporter.Base;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace DelegationExporter.Util
{
    public class HttpUtil
    {
        private static HttpClient httpClient = new HttpClient();

        public static async Task<string> Get(string url)
        {
            try
            {
                httpClient.DefaultRequestHeaders.Add("User-Agent", "DelegationExporter " + Config.VERSION);
                HttpResponseMessage data = await httpClient.GetAsync(url);
                return await data.Content.ReadAsStringAsync();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return "";
        }
    }
}
