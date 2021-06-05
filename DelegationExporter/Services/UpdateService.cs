using DelegationExporter.Base;
using DelegationExporter.Model;
using DelegationExporter.Util;
using System;
using System.Collections.Generic;

using System.Text.Json;
using System.Threading.Tasks;

namespace DelegationExporter.Services
{
    public class UpdateService
    {
        /*public async Task<bool> ChkNewVer()
        {
            try
            {
                string res = await HttpUtil.Get(Constant.GIT_RELEASE_API);
                Console.WriteLine("res:" + res);
                List<GithubReleaseRes> releaseVers = JsonSerializer.Deserialize<List<GithubReleaseRes>>(res);
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
        }*/

        public async Task<string> GetLatestVerInfo(string url)
        {
            try
            {
                string res = await HttpUtil.Get(url);
                GithubReleaseRes githubRelease = JsonSerializer.Deserialize<GithubReleaseRes>(res);
                return githubRelease.TagName;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return "";
        }
    }
}
