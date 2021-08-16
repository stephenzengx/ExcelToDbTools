using System;
using System.IO.Compression;

namespace Excel.Zip
{
    internal class ZipPackageInfo
    {
        public ZipArchiveEntry ZipArchiveEntry { get; set; }
        public string ContentType { get; set; }
        public ZipPackageInfo(ZipArchiveEntry zipArchiveEntry, string contentType)
        {
            this.ZipArchiveEntry = zipArchiveEntry;
            ContentType = contentType;
        }
    }
}
