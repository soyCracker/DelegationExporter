using System.IO.Compression;

namespace Delegation.Service.Services
{
    public class ZipService
    {
        public byte[] ZipData(Dictionary<string, byte[]> data)
        {
            using (var zipStream = new MemoryStream())
            {
                using (var zipArchive = new ZipArchive(zipStream, ZipArchiveMode.Update))
                {
                    foreach (var fileName in data.Keys)
                    {
                        var entry = zipArchive.CreateEntry(fileName);
                        using (var entryStream = entry.Open())
                        {
                            var buff = data[fileName];
                            entryStream.Write(buff, 0, buff.Length);
                        }
                    }
                }
                return zipStream.ToArray();
            }
        }
    }
}
