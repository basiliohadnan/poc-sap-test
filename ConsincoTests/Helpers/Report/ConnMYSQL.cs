//================================================================================
// Class   : ConnMYSQL
// Version : 1.01
//
// Created : 06/05/2020 - 1.00 - Gelder Carvalho - Class creation
// Updated : 21/10/2020 - 1.01 - Gelder Carvalho - Fixed connection string
//================================================================================

using System;
using System.Data;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

namespace Starline
{
    class ConnMYSQL
    {
        public MySqlConnection Conn;
        public bool Connected = false;

        public ConnMYSQL(string ConnServer, int ConnPort, string ConnDatabase, string ConnUser, string ConnPass)
        {
            try
            {
                string connectionString = "Server = " + ConnServer + "; Port = " + ConnPort.ToString() + "; Database = " + ConnDatabase + "; Uid = " + ConnUser + "; Pwd = " + ConnPass + ";";
                Conn = new MySqlConnection(connectionString);
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
    }
}
