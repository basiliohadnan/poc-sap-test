//================================================================================
// Class   : ConnPGSQL
// Version : 2.02
//
// Created : 20/04/2019 - 1.00 - Carlos Oliveira - Class creation
// Updated : 13/09/2019 - 2.00 - Gelder Carvalho - Class refactoring
// Updated : 27/01/2020 - 2.01 - Gelder Carvalho - Input Data details
// Updated : 13/02/2020 - 2.02 - Gelder Carvalho - Added Connected variable in default connection
//================================================================================

using Npgsql;
using System;
using System.IO;
using System.Data;
using System.Text;
using System.Security.Cryptography;

namespace Starline
{
    class ConnPGSQL
    {
        public NpgsqlConnection Conn;
        public bool Connected = false;

        public ConnPGSQL(string ConnServer, int ConnPort, string ConnDatabase, string ConnUser, string ConnPass)
        {
            try
            {
                string connectionString = "Server = " + ConnServer + "; Port = " + ConnPort.ToString() + "; Database = " + ConnDatabase + "; User Id = " + ConnUser + "; Password = " + ConnPass;
                Conn = new NpgsqlConnection(connectionString);
                OpenConn();
                Connected = ConnOpened();
                CloseConn();
            }
            catch
            {
                Connected = false;
            }
        }

        public ConnPGSQL(string ConnFile)
        {
            try
            {
                string ConnStr;
                ConnStr = File.ReadAllText(ConnFile);
                string connectionString = Untokenize(ConnStr);
                Conn = new NpgsqlConnection(connectionString);
                OpenConn();
                Connected = ConnOpened();
                CloseConn();
            }
            catch
            {
                Connected = false;
            }
        }

        public void OpenConn()
        {
            if (!ConnOpened())
            {
                try
                {
                    Conn.Open();
                }
                catch
                {
                }
            }
        }

        public void CloseConn()
        {
            if (ConnOpened())
            {
                Conn.Close();
            }
        }

        public bool ConnOpened()
        {
            return Conn.State == ConnectionState.Open;
        }

        private static Rijndael RijndaelInstance(string key, string initVector)
        {
            if ((key == null) || (((key.Length != 0x10) && (key.Length != 0x18)) && (key.Length != 0x20)))
            {
                return null;
            }
            if ((initVector == null) || (initVector.Length != 0x10))
            {
                return null;
            }
            Rijndael rijndael = Rijndael.Create();
            rijndael.Key = Encoding.ASCII.GetBytes(key);
            rijndael.IV = Encoding.ASCII.GetBytes(initVector);
            return rijndael;
        }

        private static byte[] HexStringToArrayBytes(string content)
        {
            try
            {
                int num = content.Length / 2;
                byte[] buffer = new byte[num];
                for (int i = 0; i < num; i++)
                {
                    buffer[i] = Convert.ToByte(content.Substring(i * 2, 2), 0x10);
                }
                return buffer;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string Untokenize(string txt)
        {
            string str2;
            if (string.IsNullOrWhiteSpace(txt))
            {
                return "00";
            }
            if ((txt.Length % 2) != 0)
            {
                return "01";
            }
            using (Rijndael rijndael = RijndaelInstance("ASDFGTRE$#@!6+|@", "QWERTYUIOP!@#+4+"))
            {
                ICryptoTransform transform = rijndael.CreateDecryptor(rijndael.Key, rijndael.IV);
                string str = null;
                try
                {
                    using (MemoryStream stream = new MemoryStream(HexStringToArrayBytes(txt)))
                    {
                        using (CryptoStream stream2 = new CryptoStream(stream, transform, CryptoStreamMode.Read))
                        {
                            using (StreamReader reader = new StreamReader(stream2))
                            {
                                try
                                {
                                    str = reader.ReadToEnd();
                                }
                                catch (Exception)
                                {
                                }
                            }
                        }
                    }
                    str2 = str;
                }
                catch (Exception)
                {
                    str2 = "";
                }
            }
            return str2;
        }

        private static string ArrayBytesToHexString(byte[] content) =>
            string.Concat(Array.ConvertAll<byte, string>(content, b => b.ToString("X2")));

        public static string Tokenize(string txt)
        {
            string str;
            if (string.IsNullOrWhiteSpace(txt))
            {
                return "00";
            }
            using (Rijndael rijndael = RijndaelInstance("ASDFGTRE$#@!6+|@", "QWERTYUIOP!@#+4+"))
            {
                ICryptoTransform transform = rijndael.CreateEncryptor(rijndael.Key, rijndael.IV);
                using (MemoryStream stream = new MemoryStream())
                {
                    using (CryptoStream stream2 = new CryptoStream(stream, transform, CryptoStreamMode.Write))
                    {
                        using (StreamWriter writer = new StreamWriter(stream2))
                        {
                            writer.Write(txt);
                        }
                    }
                    str = ArrayBytesToHexString(stream.ToArray());
                }
            }
            return str;
        }
    }
}