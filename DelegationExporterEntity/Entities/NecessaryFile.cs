using System;
using System.Collections.Generic;

namespace DelegationExporterEntity.Entities
{
    public partial class NecessaryFile
    {
        public int Id { get; set; }
        public string FileName { get; set; }
        public byte[] Data { get; set; }
        public DateTime? UpdateTime { get; set; }
    }
}
