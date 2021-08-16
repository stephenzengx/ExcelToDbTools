using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace ExcelTools
{
    /// <summary>加密算法</summary>
    public class PhoneCardNoEncryption
    {
        private static readonly SecurityOptions _securityOptions = Utils.Config.GetSection("SecurityOptions").Get<SecurityOptions>();

        /// <summary>加密</summary>
        /// <param name="str"></param>
        /// <param name="boolIv">是否使用公共初始向量</param>
        /// <returns></returns>
        public static string Encryption(string str, bool boolIv = false)
        {
            return Convert.ToBase64String(PhoneCardNoEncryption.Encode(Encoding.UTF8.GetBytes(str), boolIv));
        }

        /// <summary>加密 （是否使用公共初始向量参数）默认使用配置里的IsUseEncryptionIv</summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string Encryption(string str)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(str);
            SecurityOptions securityOptions = PhoneCardNoEncryption._securityOptions;
            int num = securityOptions != null ? (securityOptions.IsUseEncryptionIv ? 1 : 0) : 1;
            return Convert.ToBase64String(PhoneCardNoEncryption.Encode(bytes, num != 0));
        }

        /// <summary>解密</summary>
        /// <param name="str"></param>
        /// <param name="boolIv">是否使用公共初始向量</param>
        /// <returns></returns>
        public static string Decrypt(string str, bool boolIv = false)
        {
            return Encoding.UTF8.GetString(PhoneCardNoEncryption.Decode(Convert.FromBase64String(str), boolIv));
        }

        /// <summary>解密 （是否使用公共初始向量参数） 默认使用配置里的IsUseEncryptionIv</summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string Decrypt(string str)
        {
            return Encoding.UTF8.GetString(PhoneCardNoEncryption.Decode(Convert.FromBase64String(str), PhoneCardNoEncryption._securityOptions.IsUseEncryptionIv));
        }

        public static byte[] Decode(byte[] encodedData, bool boolIv = false)
        {
            using (SymmetricAlgorithm symmetricAlgorithm = PhoneCardNoEncryption.CreateSymmetricAlgorithm(boolIv))
            {
                using (HMAC hashAlgorithm = PhoneCardNoEncryption.CreateHashAlgorithm())
                {
                    byte[] numArray1 = new byte[symmetricAlgorithm.BlockSize / 8];
                    byte[] numArray2 = new byte[hashAlgorithm.HashSize / 8];
                    byte[] buffer = new byte[encodedData.Length - numArray1.Length - numArray2.Length];
                    Array.Copy((Array)encodedData, 0, (Array)numArray1, 0, numArray1.Length);
                    Array.Copy((Array)encodedData, numArray1.Length, (Array)buffer, 0, buffer.Length);
                    Array.Copy((Array)encodedData, numArray1.Length + buffer.Length, (Array)numArray2, 0, numArray2.Length);
                    if (!((IEnumerable<byte>)hashAlgorithm.ComputeHash(((IEnumerable<byte>)numArray1).Concat<byte>((IEnumerable<byte>)buffer).ToArray<byte>())).SequenceEqual<byte>((IEnumerable<byte>)numArray2))
                        throw new ArgumentException();
                    if (!boolIv)
                        symmetricAlgorithm.IV = numArray1;
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, symmetricAlgorithm.CreateDecryptor(), CryptoStreamMode.Write))
                        {
                            cryptoStream.Write(buffer, 0, buffer.Length);
                            cryptoStream.FlushFinalBlock();
                        }
                        return memoryStream.ToArray();
                    }
                }
            }
        }

        public static byte[] Encode(byte[] data, bool boolIv = false)
        {
            byte[] iv;
            byte[] array;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (SymmetricAlgorithm symmetricAlgorithm = PhoneCardNoEncryption.CreateSymmetricAlgorithm(boolIv))
                {
                    if (!boolIv)
                        symmetricAlgorithm.GenerateIV();
                    iv = symmetricAlgorithm.IV;
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, symmetricAlgorithm.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cryptoStream.Write(data, 0, data.Length);
                        cryptoStream.FlushFinalBlock();
                    }
                    array = memoryStream.ToArray();
                }
            }
            byte[] hash;
            using (HMAC hashAlgorithm = PhoneCardNoEncryption.CreateHashAlgorithm())
                hash = hashAlgorithm.ComputeHash(((IEnumerable<byte>)iv).Concat<byte>((IEnumerable<byte>)array).ToArray<byte>());
            return ((IEnumerable<byte>)iv).Concat<byte>((IEnumerable<byte>)array).Concat<byte>((IEnumerable<byte>)hash).ToArray<byte>();
        }

        public static SymmetricAlgorithm CreateSymmetricAlgorithm(bool boolIv = false)
        {
            SymmetricAlgorithm symmetricAlgorithm = SymmetricAlgorithm.Create(PhoneCardNoEncryption._securityOptions.EncryptionAlgorithm);
            symmetricAlgorithm.Key = PhoneCardNoEncryption._securityOptions.EncryptionKey.ToByteArray();
            if (boolIv)
            {
                byte[] byteArray = PhoneCardNoEncryption._securityOptions.EncryptionIv.ToByteArray();
                byte[] numArray1 = new byte[symmetricAlgorithm.BlockSize / 8];
                byte[] numArray2 = numArray1;
                int length = numArray1.Length;
                Array.Copy((Array)byteArray, 0, (Array)numArray2, 0, length);
                symmetricAlgorithm.IV = numArray1;
            }
            return symmetricAlgorithm;
        }

        public static HMAC CreateHashAlgorithm()
        {
            HMAC hmac = HMAC.Create(PhoneCardNoEncryption._securityOptions.HashAlgorithm);
            hmac.Key = PhoneCardNoEncryption._securityOptions.HashKey.ToByteArray();
            return hmac;
        }
    }
}
