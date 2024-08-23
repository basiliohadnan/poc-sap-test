//================================================================================
// Class   : InputData
// Version : 1.06
//
// Created : 27/01/2020 - 1.00 - Gelder Carvalho - Class creation (with only PostgreSQL connector)
// Updated : 12/02/2020 - 1.01 - Gelder Carvalho - Added Oracle connector
// Updated : 14/02/2020 - 1.02 - Gelder Carvalho - Bug fix - sessions static -> private
// Updated : 17/02/2020 - 1.03 - Gelder Carvalho - Added image methods for Oracle
// Updated : 18/02/2020 - 1.04 - Gelder Carvalho - Bug fix - image methods for Oracle
// Updated : 06/05/2020 - 1.05 - Gelder Carvalho - Added SQL Server connector
// Updated : 21/10/2020 - 1.06 - Gelder Carvalho - Added MySQL and Excel connector + Tests
//================================================================================

using System;
using System.Data;
using Npgsql;
using Oracle.ManagedDataAccess.Client;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Starline
{
    public class DataFetch
    {
        public bool ShowLog { get; set; }

        public string InputType { get; set; }
        public string InputServer { get; set; }
        public int InputPort { get; set; }
        public string InputUser { get; set; }
        public string InputPass { get; set; }
        public string InputDatabase { get; set; }
        public string InputService { get; set; }
        public string InputSID { get; set; }
        public string InputXLS { get; set; }
        public string InputFile { get; set; }

        private ConnPGSQL SessionPGSQL;
        private ConnORACLE SessionORACLE;
        private ConnSQLSERVER SessionSQLSERVER;
        private ConnMYSQL SessionMYSQL;
        private ConnEXCEL SessionEXCEL;

        public Dictionary<string, Dictionary<int, Dictionary<string, string>>> QueryData;

        public string Print(string IncomingText, Exception ex)
        {
            string LogMessage;
            if (ex != null)
            {
                LogMessage = IncomingText + " ( " + ex + " )";
            }
            else
            {
                LogMessage = IncomingText;
            }
            if (ShowLog)
            {
                Console.Error.WriteLine(LogMessage);
            }
            return LogMessage;
        }

        public DataFetch(string ConnType = "",
                         string ConnServer = "",
                         int ConnPort = 0,
                         string ConnUser = "",
                         string ConnPass = "",
                         string ConnDatabase = "",
                         string ConnService = "",
                         string ConnSID = "",
                         string ConnXLS = "",
                         string ConnFile = "")
        {
            ShowLog = false;
            InputType = ConnType;
            InputServer = ConnServer;
            InputPort = ConnPort;
            switch (InputType)
            {
                case "PostgreSQL":
                    //Print("PostgreSQL connector selected", null);
                    if (InputPort == 0)
                    {
                        InputPort = 5432;
                    }
                    break;

                case "Oracle":
                    //Print("Oracle connector selected", null);
                    if (InputPort == 0)
                    {
                        InputPort = 1521;
                    }
                    break;

                case "SQL Server":
                    //Print("SQL Server connector selected", null);
                    if (InputPort == 0)
                    {
                        InputPort = 1433;
                    }
                    break;

                case "MySQL":
                    //Print("MySQL connector selected", null);
                    if (InputPort == 0)
                    {
                        InputPort = 3306;
                    }
                    break;

                case "Excel":
                    //Print("Excel connector selected", null);
                    break;


                default:
                    throw new InvalidOperationException("Invalid connector type (valid options: PostgreSQL, Oracle, MySQL, Excel, SQL Server)");
            }
            InputUser = ConnUser;
            InputPass = ConnPass;
            InputDatabase = ConnDatabase;
            InputService = ConnService;
            InputSID = ConnSID;
            InputXLS = ConnXLS;
            InputFile = ConnFile;
            QueryData = new Dictionary<string, Dictionary<int, Dictionary<string, string>>>();
            try
            {
                switch (InputType)
                {
                    case "PostgreSQL":
                        Print("PostgreSQL connector selected", null);
                        if (InputFile != "")
                        {
                            SessionPGSQL = new ConnPGSQL(ConnFile);
                        }
                        else
                        {
                            SessionPGSQL = new ConnPGSQL(ConnServer, ConnPort, ConnDatabase, ConnUser, ConnPass);
                        }
                        break;

                    case "Oracle":
                        Print("Oracle connector selected", null);
                        if (ConnService != "")
                        {
                            SessionORACLE = new ConnORACLE(ConnServer, ConnPort, ConnUser, ConnPass, ConnService: ConnService);
                        }
                        if (ConnSID != "")
                        {
                            SessionORACLE = new ConnORACLE(ConnServer, ConnPort, ConnUser, ConnPass, ConnSID: ConnSID);
                        }
                        break;

                    case "SQL Server":
                        Print("SQL Server connector selected", null);
                        SessionSQLSERVER = new ConnSQLSERVER(ConnServer, ConnPort, ConnDatabase, ConnUser, ConnPass);
                        break;

                    case "MySQL":
                        Print("MySQL connector selected", null);
                        SessionMYSQL = new ConnMYSQL(ConnServer, ConnPort, ConnDatabase, ConnUser, ConnPass);
                        break;

                    case "Excel":
                        Print("Excel connector selected", null);
                        SessionEXCEL = new ConnEXCEL(ConnXLS);
                        break;

                    default:
                        throw new InvalidOperationException("Invalid connector type (valid options: PostgreSQL, Oracle, MySQL, Excel, SQL Server)");
                }
                QueryData = new Dictionary<string, Dictionary<int, Dictionary<string, string>>>();
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Error creating connector", ex));
            }
        }

        public int NewQuery(string QueryName = "", string QueryText = "", params string[] ParamsList)
        {
            int RowCount = 0;
            int CurrentRow = 0;
            try
            {
                switch (InputType)
                {
                    case "PostgreSQL":
                        if (SessionPGSQL.Connected)
                        {
                            SessionPGSQL.OpenConn();
                            if (SessionPGSQL.ConnOpened())
                            {
                                NpgsqlCommand cmdPGSQL = new NpgsqlCommand(QueryText, SessionPGSQL.Conn);
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdPGSQL.Parameters.AddWithValue(key, Convert.ToInt32(value));
                                    }
                                    else
                                    {
                                        cmdPGSQL.Parameters.AddWithValue(key, value);
                                    }
                                }
                                NpgsqlDataReader drPGSQL = cmdPGSQL.ExecuteReader();
                                if (drPGSQL.HasRows)
                                {
                                    if (QueryData.ContainsKey(QueryName))
                                    {
                                        QueryData.Remove(QueryName);
                                    }
                                    Dictionary<int, Dictionary<string, string>> RowsList = new Dictionary<int, Dictionary<string, string>>();
                                    while (drPGSQL.Read())
                                    {
                                        CurrentRow += 1;
                                        Dictionary<string, string> ColumnsList = new Dictionary<string, string>();
                                        for (int f = 0; f < drPGSQL.FieldCount; f++)
                                        {
                                            ColumnsList.Add(drPGSQL.GetName(f).ToUpper(), drPGSQL[f].ToString());
                                        }
                                        RowsList.Add(CurrentRow, ColumnsList);
                                    }
                                    QueryData.Add(QueryName, RowsList);
                                }
                                RowCount = CurrentRow;
                                SessionPGSQL.CloseConn();
                            }
                            else
                            {
                                RowCount = -1;
                                Print("Connection not established on query execution", null);
                            }
                        }
                        else
                        {
                            RowCount = -1;
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "Oracle":
                        if (SessionORACLE.Connected)
                        {
                            SessionORACLE.OpenConn();
                            if (SessionORACLE.ConnOpened())
                            {
                                OracleCommand cmdORACLE = SessionORACLE.Conn.CreateCommand();
                                cmdORACLE.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdORACLE.Parameters.Add(new OracleParameter(key, Convert.ToInt32(value)));
                                    }
                                    else
                                    {
                                        cmdORACLE.Parameters.Add(new OracleParameter(key, value));
                                    }
                                }
                                OracleDataReader drORACLE = cmdORACLE.ExecuteReader();
                                if (drORACLE.HasRows)
                                {
                                    if (QueryData.ContainsKey(QueryName))
                                    {
                                        QueryData.Remove(QueryName);
                                    }
                                    Dictionary<int, Dictionary<string, string>> RowsList = new Dictionary<int, Dictionary<string, string>>();
                                    while (drORACLE.Read())
                                    {
                                        CurrentRow += 1;
                                        Dictionary<string, string> ColumnsList = new Dictionary<string, string>();
                                        for (int f = 0; f < drORACLE.FieldCount; f++)
                                        {
                                            ColumnsList.Add(drORACLE.GetName(f).ToUpper(), drORACLE[f].ToString());
                                        }
                                        RowsList.Add(CurrentRow, ColumnsList);
                                    }
                                    QueryData.Add(QueryName, RowsList);
                                }
                                RowCount = CurrentRow;
                                SessionORACLE.CloseConn();
                            }
                            else
                            {
                                RowCount = -1;
                                Print("Connection not established on query execution", null);
                            }
                        }
                        else
                        {
                            RowCount = -1;
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "SQL Server":
                        if (SessionSQLSERVER.Connected)
                        {
                            SessionSQLSERVER.OpenConn();
                            if (SessionSQLSERVER.ConnOpened())
                            {
                                SqlCommand cmdSQLSERVER = SessionSQLSERVER.Conn.CreateCommand();
                                cmdSQLSERVER.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdSQLSERVER.Parameters.Add("@" + key, SqlDbType.Int);
                                    }
                                    else
                                    {
                                        cmdSQLSERVER.Parameters.Add("@" + key, SqlDbType.VarChar);
                                    }
                                    cmdSQLSERVER.Parameters["@" + key].Value = value;
                                }
                                SqlDataReader drSQLSERVER = cmdSQLSERVER.ExecuteReader();
                                if (drSQLSERVER.HasRows)
                                {
                                    if (QueryData.ContainsKey(QueryName))
                                    {
                                        QueryData.Remove(QueryName);
                                    }
                                    Dictionary<int, Dictionary<string, string>> RowsList = new Dictionary<int, Dictionary<string, string>>();
                                    while (drSQLSERVER.Read())
                                    {
                                        CurrentRow += 1;
                                        Dictionary<string, string> ColumnsList = new Dictionary<string, string>();
                                        for (int f = 0; f < drSQLSERVER.FieldCount; f++)
                                        {
                                            ColumnsList.Add(drSQLSERVER.GetName(f).ToUpper(), drSQLSERVER[f].ToString());
                                        }
                                        RowsList.Add(CurrentRow, ColumnsList);
                                    }
                                    QueryData.Add(QueryName, RowsList);
                                }
                                RowCount = CurrentRow;
                                SessionSQLSERVER.CloseConn();
                            }
                            else
                            {
                                RowCount = -1;
                                Print("Connection not established on query execution", null);
                            }
                        }
                        else
                        {
                            RowCount = -1;
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "MySQL":
                        if (SessionMYSQL.Connected)
                        {
                            SessionMYSQL.OpenConn();
                            if (SessionMYSQL.ConnOpened())
                            {
                                MySqlCommand cmdMYSQL = SessionMYSQL.Conn.CreateCommand();
                                cmdMYSQL.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdMYSQL.Parameters.AddWithValue("@" + key, Convert.ToInt32(value));
                                    }
                                    else
                                    {
                                        cmdMYSQL.Parameters.AddWithValue("@" + key, value);
                                    }
                                }
                                MySqlDataReader drMYSQL = cmdMYSQL.ExecuteReader();
                                if (drMYSQL.HasRows)
                                {
                                    if (QueryData.ContainsKey(QueryName))
                                    {
                                        QueryData.Remove(QueryName);
                                    }
                                    Dictionary<int, Dictionary<string, string>> RowsList = new Dictionary<int, Dictionary<string, string>>();
                                    while (drMYSQL.Read())
                                    {
                                        CurrentRow += 1;
                                        Dictionary<string, string> ColumnsList = new Dictionary<string, string>();
                                        for (int f = 0; f < drMYSQL.FieldCount; f++)
                                        {
                                            ColumnsList.Add(drMYSQL.GetName(f).ToUpper(), drMYSQL[f].ToString());
                                        }
                                        RowsList.Add(CurrentRow, ColumnsList);
                                    }
                                    QueryData.Add(QueryName, RowsList);
                                }
                                RowCount = CurrentRow;
                                SessionMYSQL.CloseConn();
                            }
                            else
                            {
                                RowCount = -1;
                                Print("Connection not established on query execution", null);
                            }
                        }
                        else
                        {
                            RowCount = -1;
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "Excel":
                        if (SessionEXCEL.Connected)
                        {
                            SessionEXCEL.OpenConn();
                            if (SessionEXCEL.ConnOpened())
                            {
                                OleDbCommand cmdEXCEL = SessionEXCEL.Conn.CreateCommand();
                                cmdEXCEL.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdEXCEL.Parameters.AddWithValue("@" + key, Convert.ToInt32(value));
                                    }
                                    else
                                    {
                                        cmdEXCEL.Parameters.AddWithValue("@" + key, value);
                                    }
                                }
                                OleDbDataReader drEXCEL = cmdEXCEL.ExecuteReader();
                                if (drEXCEL.HasRows)
                                {
                                    if (QueryData.ContainsKey(QueryName))
                                    {
                                        QueryData.Remove(QueryName);
                                    }
                                    Dictionary<int, Dictionary<string, string>> RowsList = new Dictionary<int, Dictionary<string, string>>();
                                    while (drEXCEL.Read())
                                    {
                                        CurrentRow += 1;
                                        Dictionary<string, string> ColumnsList = new Dictionary<string, string>();
                                        for (int f = 0; f < drEXCEL.FieldCount; f++)
                                        {
                                            ColumnsList.Add(drEXCEL.GetName(f).ToUpper(), drEXCEL[f].ToString());
                                        }
                                        RowsList.Add(CurrentRow, ColumnsList);
                                    }
                                    QueryData.Add(QueryName, RowsList);
                                }
                                RowCount = CurrentRow;
                                SessionEXCEL.CloseConn();
                            }
                            else
                            {
                                RowCount = -1;
                                Print("Connection not established on query execution", null);
                            }
                        }
                        else
                        {
                            RowCount = -1;
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    default:
                        throw new InvalidOperationException("Method not available for this connector type (valid options: PostgreSQL, Oracle)");
                }
                return RowCount;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Error running query", ex));
            }
        }

        public int RowCount(string QueryName = "")
        {
            if (QueryData.ContainsKey(QueryName))
            {
                return QueryData[QueryName].Count;
            }
            else
            {
                return -1;
            }
        }

        public string GetValue(string FieldName = "", string QueryName = "", int RowNumber = 1)
        {
            if (QueryData.ContainsKey(QueryName) && QueryData[QueryName].ContainsKey(RowNumber) && QueryData[QueryName][RowNumber].ContainsKey(FieldName))
            {
                return QueryData[QueryName][RowNumber][FieldName];
            }
            else
            {
                return "";
            }
        }

        public string RunDDL(string QueryText = "", bool ReturningField = false, params string[] ParamsList)
        {
            string DDLResult = "";
            try
            {
                switch (InputType)
                {
                    case "PostgreSQL":
                        if (SessionPGSQL.Connected)
                        {
                            SessionPGSQL.OpenConn();
                            if (SessionPGSQL.ConnOpened())
                            {
                                NpgsqlCommand cmdPGSQL = new NpgsqlCommand(QueryText, SessionPGSQL.Conn);
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdPGSQL.Parameters.AddWithValue(key, Convert.ToInt32(value));
                                    }
                                    else
                                    {
                                        cmdPGSQL.Parameters.AddWithValue(key, value);
                                    }
                                }
                                if (ReturningField == false)
                                {
                                    DDLResult = cmdPGSQL.ExecuteNonQuery().ToString();
                                }
                                else
                                {
                                    DDLResult = cmdPGSQL.ExecuteScalar().ToString();
                                }
                                SessionPGSQL.CloseConn();
                            }
                            else
                            {
                                DDLResult = "-1";
                                Print("Connection not established on ddl execution", null);
                            }
                        }
                        else
                        {
                            DDLResult = "-1";
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "Oracle":
                        if (SessionORACLE.Connected)
                        {
                            SessionORACLE.OpenConn();
                            if (SessionORACLE.ConnOpened())
                            {
                                OracleCommand cmdORACLE = SessionORACLE.Conn.CreateCommand();
                                cmdORACLE.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdORACLE.Parameters.Add(new OracleParameter(key, Convert.ToInt32(value)));
                                    }
                                    else
                                    {
                                        cmdORACLE.Parameters.Add(new OracleParameter(key, value));
                                    }
                                }
                                if (ReturningField == false)
                                {
                                    DDLResult = cmdORACLE.ExecuteNonQuery().ToString();
                                }
                                else
                                {
                                    DDLResult = cmdORACLE.ExecuteScalar().ToString();
                                }
                                SessionORACLE.CloseConn();
                            }
                            else
                            {
                                DDLResult = "-1";
                                Print("Connection not established on ddl execution", null);
                            }
                        }
                        else
                        {
                            DDLResult = "-1";
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "SQL Server":
                        if (SessionSQLSERVER.Connected)
                        {
                            SessionSQLSERVER.OpenConn();
                            if (SessionSQLSERVER.ConnOpened())
                            {
                                SqlCommand cmdSQLSERVER = SessionSQLSERVER.Conn.CreateCommand();
                                cmdSQLSERVER.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdSQLSERVER.Parameters.Add("@" + key, SqlDbType.Int);
                                    }
                                    else
                                    {
                                        cmdSQLSERVER.Parameters.Add("@" + key, SqlDbType.VarChar);
                                    }
                                    cmdSQLSERVER.Parameters["@" + key].Value = value;
                                }
                                if (ReturningField == false)
                                {
                                    DDLResult = cmdSQLSERVER.ExecuteNonQuery().ToString();
                                }
                                else
                                {
                                    DDLResult = cmdSQLSERVER.ExecuteScalar().ToString();
                                }
                                SessionSQLSERVER.CloseConn();
                            }
                            else
                            {
                                DDLResult = "-1";
                                Print("Connection not established on ddl execution", null);
                            }
                        }
                        else
                        {
                            DDLResult = "-1";
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "MySQL":
                        if (SessionMYSQL.Connected)
                        {
                            SessionMYSQL.OpenConn();
                            if (SessionMYSQL.ConnOpened())
                            {
                                MySqlCommand cmdMYSQL = SessionMYSQL.Conn.CreateCommand();
                                cmdMYSQL.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdMYSQL.Parameters.AddWithValue("@" + key, Convert.ToInt32(value));
                                    }
                                    else
                                    {
                                        cmdMYSQL.Parameters.AddWithValue("@" + key, value);
                                    }
                                }
                                if (ReturningField == false)
                                {
                                    DDLResult = cmdMYSQL.ExecuteNonQuery().ToString();
                                }
                                else
                                {
                                    DDLResult = cmdMYSQL.ExecuteScalar().ToString();
                                }
                                SessionMYSQL.CloseConn();
                            }
                            else
                            {
                                DDLResult = "-1";
                                Print("Connection not established on ddl execution", null);
                            }
                        }
                        else
                        {
                            DDLResult = "-1";
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    case "Excel":
                        if (SessionEXCEL.Connected)
                        {
                            SessionEXCEL.OpenConn();
                            if (SessionEXCEL.ConnOpened())
                            {
                                OleDbCommand cmdEXCEL = SessionEXCEL.Conn.CreateCommand();
                                cmdEXCEL.CommandText = QueryText;
                                for (var p = 0; p <= ParamsList.Length - 1; p++)
                                {
                                    string key = ParamsList[p].Substring(0, ParamsList[p].IndexOf(":")).Trim();
                                    string value = ParamsList[p].Remove(0, ParamsList[p].IndexOf(":") + 1).Trim();
                                    if (key.ToLower().StartsWith("v_num"))
                                    {
                                        cmdEXCEL.Parameters.AddWithValue("@" + key, Convert.ToInt32(value));
                                    }
                                    else
                                    {
                                        cmdEXCEL.Parameters.AddWithValue("@" + key, value);
                                    }
                                }
                                if (ReturningField == false)
                                {
                                    DDLResult = cmdEXCEL.ExecuteNonQuery().ToString();
                                }
                                else
                                {
                                    DDLResult = cmdEXCEL.ExecuteScalar().ToString();
                                }
                                SessionEXCEL.CloseConn();
                            }
                            else
                            {
                                DDLResult = "-1";
                                Print("Connection not established on ddl execution", null);
                            }
                        }
                        else
                        {
                            DDLResult = "-1";
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    default:
                        throw new InvalidOperationException("Method not available for this connector type (valid options: PostgreSQL, Oracle)");
                }
                return DDLResult;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Error running query", ex));
            }
        }

        public string GetImageBase64(string QueryText = "", string PreSQL = "")
        {
            string pictureResult = "";
            try
            {
                switch (InputType)
                {
                    case "Oracle":
                        if (SessionORACLE.Connected)
                        {
                            SessionORACLE.OpenConn();
                            if (SessionORACLE.ConnOpened())
                            {
                                OracleCommand cmd = SessionORACLE.Conn.CreateCommand();
                                if (PreSQL != "")
                                {
                                    cmd.CommandText = PreSQL;
                                    cmd.ExecuteNonQuery();
                                }
                                cmd.CommandText = QueryText;
                                OracleDataReader dr = cmd.ExecuteReader();
                                if (dr.HasRows)
                                {
                                    byte[] data;
                                    MemoryStream ms;
                                    while (dr.Read())
                                    {
                                        data = (byte[])dr[0];
                                        ms = new MemoryStream(data);
                                        byte[] ImgBytes = ms.ToArray();
                                        pictureResult = Convert.ToBase64String(ImgBytes);
                                    }
                                }
                                SessionORACLE.CloseConn();
                            }
                            else
                            {
                                pictureResult = "-1";
                                Print("Connection not established on ddl execution", null);
                            }
                        }
                        else
                        {
                            pictureResult = "-1";
                            Print("Connection not established on class instanciation", null);
                        }
                        break;

                    default:
                        throw new InvalidOperationException("Method not available for this connector type (valid options: Oracle)");
                }
                return pictureResult;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Error running query", ex));
            }
        }
    }
}