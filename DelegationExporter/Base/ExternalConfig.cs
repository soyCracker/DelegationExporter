using System;
using System.IO;
using DelegationExporter.Model;
using Newtonsoft.Json;

namespace DelegationExporter.Base
{
    public class ExternalConfig
    {
        public static ExternalConfigModel Get()
        {
            try
            {
                using (StreamReader sr = new StreamReader(Config.CONFIG_FILE))
                {
                    string jsonStr = sr.ReadToEnd();
                    //反序列化
                    return JsonConvert.DeserializeObject<ExternalConfigModel>(jsonStr);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return new ExternalConfigModel();
        }
    }
}
