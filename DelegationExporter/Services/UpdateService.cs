using DelegationExporter.Base;
using DelegationExporter.Model;
using DelegationExporter.Util;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace DelegationExporter.Services
{
    public class UpdateService
    {
        public async Task<bool> IsNewVer()
        {
            try
            {
                string res = await HttpUtil.Get(Constant.GIT_RELEASE_API);
                Console.WriteLine("res:" + res);
                List<GithubReleaseRes> releaseVers = JsonConvert.DeserializeObject<List<GithubReleaseRes>>(res);
                foreach (GithubReleaseRes release in releaseVers)
                {
                    Console.WriteLine("tag_name:" + release.tag_name);
                    Console.WriteLine("browser_download_url:" + release.assets[0].browser_download_url);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }
    }
}
