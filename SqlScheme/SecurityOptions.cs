using System.Collections.Generic;

namespace ExcelTools
{
    public class SecurityOptions
    {
        public string EncryptionAlgorithm { get; set; }

        public string EncryptionKey { get; set; }

        public string EncryptionIv { get; set; }

        public string HashAlgorithm { get; set; }

        public string HashKey { get; set; }

        public bool IsUseEncryptionIv { get; set; }
    }
}
