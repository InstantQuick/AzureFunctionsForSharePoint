using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;


namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Use to encrypt and decrypt text
    /// </summary>
    public class Encryption
    {
        /// <summary>
        /// Decrypts text given a password and salt
        /// </summary>
        /// <param name="cipherText">The text to decrypt</param>
        /// <param name="password">The password</param>
        /// <param name="salt">The salt</param>
        /// <seealso cref="Rijndael"/>
        /// <returns>Clear text</returns>
        public static string Decrypt(string cipherText, string password, string salt)
        {
            //Pad the salt if it is too short
            if (salt.Length < 8) salt = salt + new string('f', 8 - salt.Length);

            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            byte[] saltBytes = Encoding.ASCII.GetBytes(salt);
            Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(password, saltBytes);
            byte[] decryptedData = Decrypt(cipherBytes, pdb.GetBytes(32), pdb.GetBytes(16));
            return System.Text.Encoding.Unicode.GetString(decryptedData);
        }
        /// <summary>
        /// Encrypts text given a password and salt
        /// </summary>
        /// <param name="clearText">The text to encrypt</param>
        /// <param name="password">The password</param>
        /// <param name="salt">The salt</param>
        /// <seealso cref="Rijndael"/>
        /// <returns>Cipher text</returns>
        public static string Encrypt(string clearText, string password, string salt)
        {
            //Pad the salt if it is too short
            if (salt.Length < 8) salt = salt + new string('f', 8 - salt.Length);

            byte[] clearBytes = System.Text.Encoding.Unicode.GetBytes(clearText);
            byte[] saltBytes = Encoding.ASCII.GetBytes(salt);
            Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(password, saltBytes);
            byte[] encryptedData = Encrypt(clearBytes, pdb.GetBytes(32), pdb.GetBytes(16));
            return Convert.ToBase64String(encryptedData);
        }

        private static byte[] Decrypt(byte[] cipherData, byte[] key, byte[] IV)
        {
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = null;
            try
            {
                Rijndael alg = Rijndael.Create();
                alg.Key = key;
                alg.IV = IV;
                cs = new CryptoStream(ms, alg.CreateDecryptor(), CryptoStreamMode.Write);
                cs.Write(cipherData, 0, cipherData.Length);
                cs.FlushFinalBlock();
                return ms.ToArray();
            }
            catch
            {
                return null;
            }
            finally
            {
                cs.Close();
            }
        }

        private static byte[] Encrypt(byte[] clearData, byte[] key, byte[] IV)
        {
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = null;
            try
            {
                Rijndael alg = Rijndael.Create();
                alg.Key = key;
                alg.IV = IV;
                cs = new CryptoStream(ms, alg.CreateEncryptor(), CryptoStreamMode.Write);
                cs.Write(clearData, 0, clearData.Length);
                cs.FlushFinalBlock();
                return ms.ToArray();
            }
            catch
            {
                return null;
            }
            finally
            {
                cs.Close();
            }
        }
    }
}
