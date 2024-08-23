//================================================================================
// Class   : AutoReport
// Version : 2.04
//
// Created : 20/04/2019 - 1.00 - Carlos Oliveira - Class creation
// Updated : 13/09/2019 - 2.00 - Gelder Carvalho - Class refactoring
// Updated : 30/09/2019 - 2.01 - Gelder Carvalho - ReportProfile parameter
// Updated : 07/10/2019 - 2.02 - Gelder Carvalho - Default ReportTitle
// Updated : 24/01/2020 - 2.03 - Gelder Carvalho - Default Responsible Analyst
// Updated : 30/04/2020 - 2.04 - Gelder Carvalho - Added obtainedResult in EndStep function
//================================================================================

using Npgsql;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using Newtonsoft.Json;
using SAPTests.Helpers;

namespace Starline
{
    public class AutoReport
    {
        public bool mostrarLog { get; set; }

        static ConnPGSQL Conn;
        public bool Connected = false;
        public string URLDomain { get; set; }
        public string ResponsibleAnalyst { get; set; }
        public int CstID { get; set; }
        public string CustomerName { get; set; }
        public int SitID { get; set; }
        public string SuiteName { get; set; }
        public int ScnID { get; set; }
        public string ScenarioName { get; set; }
        public int TstID { get; set; }
        public string TestName { get; set; }
        public string TestType { get; set; }
        public string TestDesc { get; set; }
        public string AnalystName { get; set; }
        public string PreCondition { get; set; }
        public string PostCondition { get; set; }
        public string InputData { get; set; }
        public int RptID { get; set; }
        public int ReportID { get; set; }
        public string ReportTitle { get; set; }
        public string ReportObs { get; set; }
        public string ReportProfile { get; set; }
        public int LgsID { get; set; }
        public string StepName { get; set; }
        public int StepNumber { get; set; }
        public int StepTurn { get; set; }
        public Dictionary<string, Dictionary<string, int>> TurnList;

        public AutoReport()
        {
            mostrarLog = true;
            ReportTitle = "Evolução dos Casos de Teste Automatizados" + " - " + DateTime.Now.ToString("dd'/'MM'/'yyyy");

            string appPath = GetAppPath();
            string relativePathToFile = Path.Combine("..", "..", "..", "..", "ConsincoTests", "secrets.json");
            string secretsFilePath = FileHelper.GetFullPathFromBase(relativePathToFile);
            string secretsJson = File.ReadAllText(secretsFilePath);
            dynamic secrets = JsonConvert.DeserializeObject(secretsJson);
            string ConnServer = secrets.ConnServer;
            int ConnPort = secrets.ConnPort;
            string ConnDatabase = secrets.ConnDatabase;
            string ConnUser = secrets.ConnUser;
            string ConnPass = secrets.ConnPass;

            Conn = new ConnPGSQL(ConnServer, ConnPort, ConnDatabase, ConnUser, ConnPass);
            try
            {
                Conn.OpenConn();
                Connected = Conn.ConnOpened();
                Conn.CloseConn();
            }
            catch
            {
                Connected = false;
            }

            TurnList = new Dictionary<string, Dictionary<string, int>>();
        }

        public string GetAppPath()
        {
            try
            {
                string getDir = AppDomain.CurrentDomain.BaseDirectory.Replace("\\", "/");
                string projectDir;
                string debugDir;
                string binDir;
                string appDir;

                var targetFrameWork = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(System.Runtime.Versioning.TargetFrameworkAttribute), false);
                var frameWork = ((System.Runtime.Versioning.TargetFrameworkAttribute)(targetFrameWork[0])).FrameworkName;
                string classFrameWork = frameWork.Substring(0, frameWork.IndexOf(","));

                if (classFrameWork == ".NETCoreApp")
                {
                    projectDir = getDir.Substring(0, getDir.LastIndexOf("/"));
                    debugDir = projectDir.Substring(0, projectDir.LastIndexOf("/"));
                    binDir = debugDir.Substring(0, debugDir.LastIndexOf("/"));
                    appDir = binDir.Substring(0, binDir.LastIndexOf("/"));
                }
                else
                {
                    debugDir = getDir.Substring(0, getDir.LastIndexOf("/"));
                    binDir = debugDir.Substring(0, debugDir.LastIndexOf("/"));
                    appDir = binDir.Substring(0, binDir.LastIndexOf("/"));
                }
                return appDir;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at GetAppPath", ex));
            }
        }

        public string Print(string methodName, Exception ex)
        {
            string msg;
            if (ex != null)
            {
                msg = methodName + " ( " + ex + " )";
            }
            else
            {
                msg = methodName;
            }
            if (mostrarLog)
            {
                Console.Error.WriteLine(msg);
            }
            return msg;
        }

        public string Call(string functionName, params string[] paramsList)
        {
            if (Connected)
            {
                Conn.OpenConn();
                if (Conn.ConnOpened())
                {
                    try
                    {
                        string select = "select " + functionName + "(";
                        for (var x = 0; x <= paramsList.Length - 1; x++)
                        {
                            if (x == 0)
                            {
                                select += "@CP" + x.ToString();
                            }
                            else
                            {
                                select += ",@CP" + x.ToString();
                            }
                        }
                        select += ") as ID";

                        NpgsqlCommand cmd = new NpgsqlCommand(select, Conn.Conn);
                        for (var x = 0; x <= paramsList.Length - 1; x++)
                        {
                            string key = paramsList[x].Substring(0, paramsList[x].IndexOf(":"));
                            string value = paramsList[x].Remove(0, paramsList[x].IndexOf(":") + 1);
                            if (key.ToLower().IndexOf("_id") > 0)
                            {
                                cmd.Parameters.AddWithValue("CP" + x.ToString(), Convert.ToInt32(value));
                            }
                            else
                            {
                                cmd.Parameters.AddWithValue("CP" + x.ToString(), value);
                            }
                        }
                        string result = "";
                        NpgsqlDataReader dr = cmd.ExecuteReader();
                        if (dr.HasRows && dr.Read())
                        {
                            result = dr["ID"].ToString();
                        }
                        Conn.CloseConn();
                        return result;
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(Print("Exception at Call", ex));
                    }
                }
                else
                {
                    return "1";
                }
            }
            else
            {
                return "1";
            }
        }

        public string Slugify(string value)
        {
            try
            {
                //LowerCase
                value = value.ToLowerInvariant();

                //RemoveAccents
                Byte[] bytes = Encoding.GetEncoding("us-ascii").GetBytes(value);
                value = Encoding.UTF8.GetString(bytes);

                //RemoveSpaces
                value = Regex.Replace(value, @"\s", "-", RegexOptions.Compiled);

                //RemoveSpecialCharacters
                value = Regex.Replace(value, @"[^\w\s\p{Pd}]", "", RegexOptions.Compiled);

                //RemoveFinalDashes
                value = value.Trim('-', '_');

                //RemoveDoubleDashes
                value = Regex.Replace(value, @"([-_]){2,}", "$1", RegexOptions.Compiled);

                //Slugified
                return value;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at Slugify", ex));
            }
        }

        public int StartTest(string customerName, string suiteName, string scenarioName, string testName, string testType, string analystName, string testDesc, int reportID = 0, int deleteFlag = 1, string statusType = "worst")
        {
            try
            {
                if (Connected == false)
                {
                    Print("--------------------------------------------------------------------------------", null);
                    Print("  Report Database Offline", null);
                }
                Print("--------------------------------------------------------------------------------", null);
                CustomerName = Slugify(customerName);
                CstID = Convert.ToInt32(Call("arch.update_cst", "customer:" + CustomerName));
                Print("  Customer ID: " + CstID.ToString() + " [" + CustomerName + "]", null);
                SuiteName = Slugify(suiteName);
                SitID = Convert.ToInt32(Call("auto.update_sit", "cst_id:" + CstID.ToString(), "suite:" + SuiteName));
                Print("     Suite ID: " + SitID.ToString() + " [" + SuiteName + "]", null);
                ScenarioName = Slugify(scenarioName);
                ScnID = Convert.ToInt32(Call("auto.update_scn", "cst_id:" + CstID.ToString(), "scenario:" + ScenarioName));
                Print("  Scenario ID: " + ScnID.ToString() + " [" + ScenarioName + "]", null);
                TestName = Slugify(testName);
                TestType = testType;
                TestDesc = testDesc;
                AnalystName = analystName;
                TstID = Convert.ToInt32(Call("auto.update_tst", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "test:" + TestName, "type:" + TestType, "desc:" + TestDesc, "analyst:" + AnalystName, "status:" + statusType));
                Print("      Test ID: " + TstID.ToString() + " [" + TestName + "]", null);
                Print("--------------------------------------------------------------------------------", null);
                if (deleteFlag == 1)
                {
                    if (ReportID != 0)
                    {
                        Call("auto.delete_rpt", "cst_id:" + CstID.ToString(), "rpt_id:" + reportID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString());
                        Call("auto.delete_lgs", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + reportID.ToString());
                    }
                    Call("auto.delete_stp", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString());
                }
                RptID = Convert.ToInt32(Call("auto.update_rpt", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + reportID.ToString(), "status:pendente"));
                Print("    Report ID: " + RptID.ToString(), null);
                Print("--------------------------------------------------------------------------------", null);
                return RptID;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at StartTest", ex));
            }
        }

        public void DoTest(string pre = "", string post = "", string inputData = "", int steps = 0)
        {
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("DoTest [CstID=0]", null));
                }
                if (SitID == 0)
                {
                    throw new Exception(Print("DoTest [SitID=0]", null));
                }
                if (ScnID == 0)
                {
                    throw new Exception(Print("DoTest [ScnID=0]", null));
                }
                if (TstID == 0)
                {
                    throw new Exception(Print("DoTest [TstID=0]", null));
                }
                Call("auto.update_tst_details", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id: " + TstID.ToString(), "pre:" + pre, "post:" + post, "input_data:" + inputData, "steps_id:" + steps.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at DoTest", ex));
            }

        }

        public int DoStep(string stepDesc, string expectedResult = "", string status = "active", int newStep = 1)
        {
            expectedResult = $"Sucesso ao {stepDesc.ToLower()}";
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("DoStep [CstID=0]", null));
                }
                if (SitID == 0)
                {
                    throw new Exception(Print("DoStep [SitID=0]", null));
                }
                if (ScnID == 0)
                {
                    throw new Exception(Print("DoStep [ScnID=0]", null));
                }
                if (TstID == 0)
                {
                    throw new Exception(Print("DoStep [TstID=0]", null));
                }
                if (RptID == 0)
                {
                    throw new Exception(Print("DoStep [RptID=0]", null));
                }
                if (Connected)
                {
                    StepNumber = Convert.ToInt32(Call("auto.update_stp", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString(), "dsc:" + stepDesc, "expected_result:" + expectedResult, "status:" + status));
                }
                else
                {
                    StepNumber = LocalTurn(TestName + "_DoStep", stepDesc + "_DoStep");
                }
                if (newStep == 1)
                {
                    Call("auto.delete_lgs", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + RptID.ToString(), "description:" + stepDesc);
                    StartStep(stepDesc, 1, status = "pendente");
                }
                Call("auto.count_stp", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString());
                return StepNumber;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at DoStep", ex));
            }
        }

        // Compatibility with older codes
        public int NovoTurno(string stringOne, string stringTwo)
        {
            try
            {
                TurnList[stringOne][stringTwo] += 1;
                return TurnList[stringOne][stringTwo];
            }
            catch
            {
                if (TurnList.ContainsKey(stringOne))
                {
                    TurnList[stringOne].Add(stringTwo, 1);
                }
                else
                {
                    TurnList.Add(stringOne, new Dictionary<string, int>() { { stringTwo, 1 } });
                }
                return TurnList[stringOne][stringTwo];
            }
        }

        public int LocalTurn(string stringOne, string stringTwo, string mode = "add")
        {
            try
            {
                int newTurn = 0;
                try
                {
                    if (mode == "add")
                    {
                        TurnList[stringOne][stringTwo] += 1;
                    }
                    newTurn = TurnList[stringOne][stringTwo];
                }
                catch
                {
                    if (TurnList.ContainsKey(stringOne))
                    {
                        TurnList[stringOne].Add(stringTwo, 1);
                        TurnList[stringOne].Add(stringTwo + "_index", (TurnList[stringOne].Count() + 1) / 2);
                    }
                    else
                    {
                        TurnList.Add(stringOne, new Dictionary<string, int>() { { stringTwo, 1 } });
                        TurnList[stringOne].Add(stringTwo + "_index", (TurnList[stringOne].Count() + 1) / 2);
                    }
                }
                if (mode == "index")
                {
                    return TurnList[stringOne][stringTwo + "_index"];
                }
                newTurn = TurnList[stringOne][stringTwo];
                return newTurn;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at LocalTurn", ex));
            }
        }

        public int NextTurn(int stepNumber, string mode = "add")
        {
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("NextTurn [CstID=0]", null));
                }
                if (RptID == 0)
                {
                    throw new Exception(Print("NextTurn [RptID=0]", null));
                }
                if (SitID == 0)
                {
                    throw new Exception(Print("NextTurn [SitID=0]", null));
                }
                if (ScnID == 0)
                {
                    throw new Exception(Print("NextTurn [ScnID=0]", null));
                }
                if (TstID == 0)
                {
                    throw new Exception(Print("NextTurn [TstID=0]", null));
                }
                if (Connected)
                {
                    StepTurn = Convert.ToInt32(Call("auto.get_trn_num", "cst_id:" + CstID.ToString(), "rpt_id:" + RptID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "step_id:" + stepNumber.ToString()));
                    if (mode == "add")
                    {
                        StepTurn += 1;
                    }
                }
                else
                {
                    StepTurn = LocalTurn(TestName, StepName, mode);
                }
                return StepTurn;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at NextTurn", ex));
            }
        }

        public int StartStep(string stepDesc, int turn = 0, string status = "executando", string logMsg = "", string paramName = "", string paramValue = "")
        {
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("StartStep [CstID=0]", null));
                }
                if (SitID == 0)
                {
                    throw new Exception(Print("StartStep [SitID=0]", null));
                }
                if (ScnID == 0)
                {
                    throw new Exception(Print("StartStep [ScnID=0]", null));
                }
                if (TstID == 0)
                {
                    throw new Exception(Print("StartStep [TstID=0]", null));
                }
                if (RptID == 0)
                {
                    throw new Exception(Print("StartStep [RptID=0]", null));
                }
                StepName = stepDesc;
                if (turn == 0)
                {
                    StepTurn = NextTurn(GetStepNumber(StepName));
                }
                else
                {
                    StepTurn = turn;
                }
                if (status == "executando")
                {
                    string logLine = "  Step " + GetStepNumber(StepName).ToString() + "." + StepTurn.ToString() + ": " + StepName;
                    if (logMsg != "" || paramName != "")
                    {
                        logLine = logLine + " (";
                        if (logMsg != "")
                        {
                            logLine = logLine + "Obs: " + logMsg;
                            if (paramName != "")
                            {
                                logLine = logLine + " - ";
                            }
                        }
                        if (paramName != "")
                        {
                            logLine = logLine + "Parameter: " + paramName + " = " + paramValue;
                        }
                        logLine = logLine + ")";
                    }
                    Print(logLine, null);
                }
                return Convert.ToInt32(Call("auto.update_lgs", "cst_id:" + CstID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString(), "rpt_id:" + RptID.ToString(), "dsc:" + StepName, "turn_id:" + StepTurn.ToString(), "status:" + status, "log_msg:" + logMsg, "param_name:" + paramName, "param_value:" + paramValue));
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at StartStep", ex));
            }
        }

        public int GetStepNumber(string stepName)
        {
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("GetStepNumber [CstID=0]", null));
                }
                if (TstID == 0)
                {
                    throw new Exception(Print("GetStepNumber [TstID=0]", null));
                }
                if (Connected)
                {
                    return Convert.ToInt32(Call("auto.get_stp_num", "cst_id:" + CstID.ToString(), "tst_id:" + TstID.ToString(), "desc:" + stepName));
                }
                else
                {
                    return LocalTurn(TestName + "_DoStep", stepName + "_DoStep", "index");
                }

            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at GetStepNumber", ex));
            }
        }

        public byte[] ImageToByte(Image printScreen)
        {
            try
            {
                if (printScreen != null)
                {
                    MemoryStream ms = new MemoryStream();
                    printScreen.Save(ms, ImageFormat.Png);
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] imgByte = new byte[ms.Length];
                    ms.Read(imgByte, 0, Convert.ToInt32(ms.Length));
                    return imgByte;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception(Print("Exception at ImageToByte", ex));
            }
        }

        public void PutImage(int CstId, int RptId, int lgsID, string printPath)
        {
            if (Connected)
            {
                try
                {
                    Conn.OpenConn();
                    if (Conn.ConnOpened())
                    {
                        NpgsqlCommand cmd = new NpgsqlCommand("select auto.put_image(@CstId,@RptId,@LogsId,@ImageData) as ID", Conn.Conn);
                        cmd.Parameters.AddWithValue("ImageData", ImageToByte(Image.FromFile(printPath)));
                        cmd.Parameters.AddWithValue("LogsId", lgsID);
                        cmd.Parameters.AddWithValue("CstId", CstId);
                        cmd.Parameters.AddWithValue("RptId", RptId);
                        cmd.ExecuteNonQuery();
                    }
                    Conn.CloseConn();
                }
                catch (Exception ex)
                {
                    Print("Exception at PutImage", ex);
                }
            }
        }

        public bool EndStep(int lgsID, string status = "sucesso", string printPath = "", string logMsg = "", string reason = "", string obtainedResult = "")
        {
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("EndStep [CstID=0]", null));
                }
                if (RptID == 0)
                {
                    throw new Exception(Print("EndStep [RptID=0]", null));
                }
                if (logMsg != "")
                {
                    if (status == "sucesso")
                    {
                        Print("    Obs: " + logMsg, null);
                    }
                    else
                    {
                        Print("    [" + status.ToUpperInvariant() + "]: " + logMsg, null);
                    }
                }
                printPath = printPath.Replace(@"\", @"/");
                Call("auto.update_lgs_details", "cst_id:" + CstID.ToString(), "lgs_id:" + lgsID.ToString(), "status:" + status, "print_path:" + printPath, "log_msg:" + logMsg, "reason:" + reason, "obtained_result:" + obtainedResult);
                if (printPath != "")
                {
                    PutImage(CstID, RptID, lgsID, printPath);
                }
                if (status == "erro" || status == "grave")
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
                throw new Exception(Print("Exception at EndStep", ex));
            }
        }

        public void EndTest(int reportID = 0, string status = "pronto", bool deleteReport = false)
        {
            try
            {
                if (CstID == 0)
                {
                    throw new Exception(Print("EndTest [CstID=0]", null));
                }
                if (RptID == 0)
                {
                    throw new Exception(Print("EndTest [RptID=0]", null));
                }
                if (SitID == 0)
                {
                    throw new Exception(Print("EndTest [SitID=0]", null));
                }
                if (ScnID == 0)
                {
                    throw new Exception(Print("EndTest [ScnID=0]", null));
                }
                if (TstID == 0)
                {
                    throw new Exception(Print("EndTest [TstID=0]", null));
                }
                if (deleteReport == true || reportID == 0)
                {
                    Call("auto.delete_rpt", "cst_id:" + CstID.ToString(), "rpt_id:" + RptID.ToString(), "sit_id:" + SitID.ToString(), "scn_id:" + ScnID.ToString(), "tst_id:" + TstID.ToString());
                }
                Print("--------------------------------------------------------------------------------", null);
                GetReportURL();
                Print("--------------------------------------------------------------------------------", null);
                if (Conn.ConnOpened())
                {
                    Conn.CloseConn();
                }
            }
            catch (Exception ex)
            {
                Print("Exception at EndTest", ex);
            }
        }

        public void GetReportURL()
        {
            try
            {
                if (String.IsNullOrEmpty(CustomerName))
                {
                    throw new Exception(Print("GetReportURL [CustomerName='']", null));
                }
                if (RptID == 0)
                {
                    throw new Exception(Print("GetReportURL [RptID=0]", null));
                }
                if (Connected)
                {
                    if (String.IsNullOrEmpty(ReportTitle))
                    {
                        throw new Exception(Print("GetReportURL [ReportTitle='']", null));
                    }
                    if (String.IsNullOrEmpty(AnalystName))
                    {
                        throw new Exception(Print("GetReportURL [AnalystName='']", null));
                    }
                    Conn.OpenConn();
                    if (Conn.ConnOpened())
                    {
                        ResponsibleAnalyst = Call("auto.get_conf", "cst_name:" + CustomerName, "conf_name:" + "default_analyst_name");
                        URLDomain = Call("auto.get_conf", "cst_name:" + CustomerName, "conf_name:" + "export_url_domain");
                        string profileParam = "";
                        if (ReportProfile != "")
                        {
                            profileParam = "&profile=" + ReportProfile;
                        }
                        string CompileURL = URLDomain + "/portal/auto/reports/update.php?customer_name=" + CustomerName + "&rpt_id=" + RptID.ToString() + "&rpt_title=" + ReportTitle + "&rpt_obs=" + ReportObs + "&rpt_analyst=" + ResponsibleAnalyst + profileParam;

                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(CompileURL);
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        string URLReturn = new StreamReader(response.GetResponseStream()).ReadToEnd();
                        if (URLReturn.Contains("[Done]"))
                        {
                            Print("  Relatório gerado com sucesso! Clique no link abaixo:", null);
                            Print("    " + URLDomain + "/portal/auto/reports/" + CustomerName + "/" + RptID.ToString(), null);
                        }
                        else
                        {
                            Print("  Relatório gerado com erro!", null);
                        }
                    }
                    Conn.CloseConn();
                }
                else
                {
                    Print("  Teste executado, porém o relatório não foi atualizado. Clique no link abaixo para acessar o último relatório disponível:", null);
                    Print("    " + URLDomain + "/portal/auto/reports/" + CustomerName + "/" + ReportID.ToString(), null);
                }
            }
            catch (Exception ex)
            {
                Print("Exception at GetReportURL", ex);
            }
        }
    }
}
