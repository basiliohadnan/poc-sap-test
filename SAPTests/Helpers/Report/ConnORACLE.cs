//================================================================================
// Class   : ConnORACLE
// Version : 1.01
//
// Created : 12/02/2020 - 1.00 - Gelder Carvalho - Class creation
// Updated : 21/10/2020 - 1.01 - Gelder Carvalho - Removed environment variables - Oracle.ManagedDataAccess.Core is not affected
//================================================================================

using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;

namespace Starline
{
    class ConnORACLE
    {
        public OracleConnection Conn;
        public bool Connected = false;

        public ConnORACLE(string ConnServer, int ConnPort, string ConnUser, string ConnPass, string ConnService = "", string ConnSID = "")
        {
            try
            {
                string instanceType = "";
                if (ConnService != "")
                {
                    instanceType = "(SERVICE_NAME = " + ConnService + ")";
                }
                if (ConnSID != "")
                {
                    instanceType = "(ORACLE_SID = " + ConnSID + ")";
                }

                string databaseTNS = "(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = " + ConnServer + ")(PORT = " + ConnPort.ToString() + "))(CONNECT_DATA = (SERVER = DEDICATED)" + instanceType + "))";
                string connectionString = "Data Source = " + databaseTNS + "; User Id = " + ConnUser + "; Password = " + ConnPass;
                Conn = new OracleConnection(connectionString);
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
