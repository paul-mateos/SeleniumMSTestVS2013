using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace RTA.Automation.CRM
{
    public class Crypto : IDisposable
    {
        private readonly Rijndael _rijndael = Rijndael.Create();
        private readonly UTF8Encoding _encoding = new UTF8Encoding();

        public Crypto(string key, string vector)
        {
            _rijndael.Key = StrToByteArray(key);
            _rijndael.IV = StrToByteArray(vector);
        }

        public Crypto(byte[] key, byte[] vector)
        {
            _rijndael.Key = key;
            _rijndael.IV = vector;
        }

        public byte[] Encrypt(string valueToEncrypt)
        {
            var bytes = _encoding.GetBytes(valueToEncrypt);

            using (var encryptor = _rijndael.CreateEncryptor())
            using (var stream = new MemoryStream())
            using (var crypto = new CryptoStream(stream, encryptor, CryptoStreamMode.Write))
            {
                crypto.Write(bytes, 0, bytes.Length);
                crypto.FlushFinalBlock();
                stream.Position = 0;
                var encrypted = new byte[stream.Length];
                stream.Read(encrypted, 0, encrypted.Length);
                return encrypted;
            }
        }

        public string Decrypt(byte[] encryptedValue)
        {
            using (var decryptor = _rijndael.CreateDecryptor())
            using (var stream = new MemoryStream())
            using (var crypto = new CryptoStream(stream, decryptor, CryptoStreamMode.Write))
            {
                crypto.Write(encryptedValue, 0, encryptedValue.Length);
                crypto.FlushFinalBlock();
                stream.Position = 0;
                var decryptedBytes = new Byte[stream.Length];
                stream.Read(decryptedBytes, 0, decryptedBytes.Length);
                return _encoding.GetString(decryptedBytes);
            }
        }

        public static byte[] GenerateEncryptionKey()
        {
            var rm = new RijndaelManaged();
            rm.GenerateKey();
            return rm.Key;
        }

        public static byte[] GenerateEncryptionVector()
        {
            var rm = new RijndaelManaged();
            rm.GenerateIV();
            return rm.IV;
        }

        public static byte[] StrToByteArray(string str)
        {
            if (str.Length == 0)
                throw new Exception("Invalid string value in StrToByteArray");

            var byteArr = new byte[str.Length / 3];
            var i = 0;
            var j = 0;

            do
            {
                var val = byte.Parse(str.Substring(i, 3));
                byteArr[j++] = val;
                i += 3;
            }

            while (i < str.Length);
            return byteArr;
        }

        public static string ByteArrToString(byte[] byteArr)
        {
            var tempStr = "";

            for (var i = 0; i <= byteArr.GetUpperBound(0); i++)
            {
                var val = byteArr[i];

                if (val < 10)
                    tempStr += "00" + val;
                else if (val < 100)
                    tempStr += "0" + val;
                else
                    tempStr += val.ToString(CultureInfo.InvariantCulture);
            }

            return tempStr;
        }

        public void Dispose()
        {
            if (_rijndael != null)
            {
                _rijndael.Dispose();
            }
        }
    }
}
