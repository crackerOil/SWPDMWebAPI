using System;
using System.Collections;
using System.Collections.Generic;

namespace SWPDMWebAPI.Containers
{
    public record PDMFile 
    {
        public string Name { get; init; }
        public bool CheckedOut { get; init; }
        public string CheckedOutBy { get; set; }
        public VersionInfo Version { get; set; }
        public byte[] Thumbnail { get; set; }
    }
}
