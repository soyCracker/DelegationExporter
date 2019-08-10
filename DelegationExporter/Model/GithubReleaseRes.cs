using System;
using System.Collections.Generic;
using System.Text;

namespace DelegationExporter.Model
{
    public class GithubReleaseRes
    {
        public string tag_name { get; set; }

        public List<GithubReleaseAsset> assets { get; set; }


    }
}
