//================================================================================
// Class   : ConnSQLSERVER
// Version : 1.00
//
// Created : 06/05/2020 - 1.00 - Gelder Carvalho - Class creation
//================================================================================

using System;
using System.Data;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;

namespace Starline
{
    class ConnSQLSERVER
    {
        public SqlConnection Conn;
        public bool Connected = false;

        public ConnSQLSERVER(string ConnServer, int ConnPort, string ConnDatabase, string ConnUser, string ConnPass)
        {
            try
            {
                string connectionString = "Server = " + ConnServer + "," + ConnPort.ToString() + "; Database = " + ConnDatabase + "; User Id = " + ConnUser + "; Password = " + ConnPass + ";";
                Conn = new SqlConnection(connectionString);
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


