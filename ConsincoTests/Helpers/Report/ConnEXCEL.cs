//================================================================================
// Class   : ConnEXCEL
// Version : 1.00
//
// Created : 21/10/2020 - 1.00 - Gelder Carvalho - Class creation
//================================================================================

using System.Data;
using System.Data.OleDb;

namespace Starline
{
    class ConnEXCEL
    {
        public OleDbConnection Conn;
        public bool Connected = false;

        public ConnEXCEL(string ConnXLS)
        {
            try
            {
                string connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ConnXLS + "; Extended Properties = 'Excel 12.0 Xml; HDR = YES'";
                Conn = new OleDbConnection(connectionString);
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
                catch (Exception ex)
                {
                    throw new Exception("ConnEXCEL", ex);
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
