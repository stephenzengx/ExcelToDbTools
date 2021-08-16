using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Excel.Zip
{
    internal class ExcelOpenXmlZip : IDisposable
    {
        private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings()
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            XmlResolver = (XmlResolver)null
        };
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private bool _disposed;
        private Stream _zipStream;
        internal MiniExcelZipArchive ZipFile;
        public ReadOnlyCollection<ZipArchiveEntry> entries;

        public ExcelOpenXmlZip(
          Stream fileStream,
          ZipArchiveMode mode = ZipArchiveMode.Read,
          bool leaveOpen = false,
          Encoding entryNameEncoding = null)
        {
            Stream stream = fileStream;
            if (stream == null)
                throw new ArgumentNullException(nameof(fileStream));
            this._zipStream = stream;
            this.ZipFile = new MiniExcelZipArchive(fileStream, mode, leaveOpen, entryNameEncoding);
            this._entries = new Dictionary<string, ZipArchiveEntry>((IEqualityComparer<string>)StringComparer.OrdinalIgnoreCase);
            this.entries = this.ZipFile.Entries;
            foreach (ZipArchiveEntry entry in this.ZipFile.Entries)
                this._entries.Add(entry.FullName.Replace('\\', '/'), entry);
        }

        public ZipArchiveEntry GetEntry(string path)
        {
            ZipArchiveEntry zipArchiveEntry;
            if (this._entries.TryGetValue(path, out zipArchiveEntry))
                return zipArchiveEntry;
            return (ZipArchiveEntry)null;
        }

        public XmlReader GetXmlReader(string path)
        {
            ZipArchiveEntry entry = this.GetEntry(path);
            if (entry != null)
                return XmlReader.Create(entry.Open(), ExcelOpenXmlZip.XmlSettings);
            return (XmlReader)null;
        }

        ~ExcelOpenXmlZip()
        {
            this.Dispose(false);
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize((object)this);
        }

        private void Dispose(bool disposing)
        {
            if (this._disposed)
                return;
            if (disposing)
            {
                if (this.ZipFile != null)
                {
                    this.ZipFile.Dispose();
                    this.ZipFile = (MiniExcelZipArchive)null;
                }
                if (this._zipStream != null)
                {
                    this._zipStream.Dispose();
                    this._zipStream = (Stream)null;
                }
            }
            this._disposed = true;
        }
    }
}