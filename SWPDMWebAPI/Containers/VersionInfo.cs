using System;

namespace SWPDMWebAPI.Containers
{
    public record VersionInfo
    {
        public int VersionNo { get; set; }
        public DateTime DateModified { get; set; }
        public string User { get; set; }
        public string Comment { get; set; }
    }
}