using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Serialization;

namespace DelegationExporter.Model
{
    public class GithubReleaseRes
    {
        [JsonPropertyName("tag_name")]
        public string TagName { get; set; }
        [JsonPropertyName("assets")]
        public List<GithubReleaseAsset> Assets { get; set; }
        public string DownloadUrl { get; set; }
        public DateTime CreatedAt { get; set; }
        public string CompressFileName { get; set; }
    }

    public class GithubReleaseAsset
    {
        [JsonPropertyName("browser_download_url")]
        public string AssetBrowserDownloadUrl { get; set; }
        [JsonPropertyName("created_at")]
        public DateTime AssetCreatedAt { get; set; }
        [JsonPropertyName("name")]
        public string AssetName { get; set; }
    }
}
