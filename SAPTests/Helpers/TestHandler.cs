using Starline;
using System.Reflection;

namespace SAPTests.Helpers
{
    public class TestHandler
    {
        public static string GetCurrentMethodName()
        {
            MethodBase method = new System.Diagnostics.StackTrace().GetFrame(1).GetMethod();
            return method.Name;
        }

        public static string SetExecutionControl(string dataFilePath, string sheet, string testName)
        {
            string executionType;
            try
            {
                executionType = Environment.GetEnvironmentVariable("executionType").ToString();
            }
            catch
            {
                executionType = "2";
                ExcelHelper.UpdateTestResult(dataFilePath, sheet, testName, "");
            }

            if (Global.firstRun)
            {
                if (executionType == "1")
                {
                    ExcelHelper.ResetTestResults(dataFilePath);
                }
                Global.firstRun = false;
            }

            Global.dataFetch = new DataFetch(ConnType: "Excel", ConnXLS: dataFilePath);
            Global.dataFetch.NewQuery(
                QueryName: testName,
                QueryText: $"SELECT * FROM [{sheet}$] WHERE testName = {testName}"
            );
            string result = Global.dataFetch.GetValue("RESULT", testName);
            return result;
        }

        public static void StartTest(DataFetch dataFetch, string testName)
        {
            string scenarioName = dataFetch.GetValue("SCENARIONAME", testName);
            string testType = dataFetch.GetValue("TESTTYPE", testName);
            string analystName = dataFetch.GetValue("ANALYSTNAME", testName);
            string testDesc = dataFetch.GetValue("TESTDESC", testName);
            int reportID;
            try
            {
                reportID = int.Parse(Environment.GetEnvironmentVariable("reportID"));
                Console.WriteLine(reportID);
            }
            catch
            {
                reportID = int.Parse(dataFetch.GetValue("REPORTID", testName));
            }
            Global.processTest.StartTest(Global.customerName, Global.suiteName, scenarioName, testName, testType, analystName, testDesc, reportID);
        }

        public static void DoTest(DataFetch dataFetch, string testName)
        {
            string preCondition = dataFetch.GetValue("PRECONDITION", testName);
            string postCondition = dataFetch.GetValue("POSTCONDITION", testName);
            string inputData = dataFetch.GetValue("INPUTDATA", testName);
            Global.processTest.DoTest(preCondition, postCondition, inputData);
        }

        public static void EndTest(DataFetch dataFetch, string testName, bool closeWindow = true)
        {
            int reportID = int.Parse(dataFetch.GetValue("REPORTID", testName));
            Global.processTest.EndTest(reportID, testName);
            if (closeWindow == true)
            {
                WindowHandler.CloseWindow();
            }
        }

        public static void DefineSteps(string testName)
        {
            switch (testName)
            {
                //App Login Screen
                case "RealizarLogin":
                    Global.processTest.DoStep("Abrir app");
                    Global.processTest.DoStep("Realizar login do analista");
                    Global.processTest.DoStep("Validar tela principal exibida");
                    break;

                default:
                    throw new Exception($"{testName} has no steps definition.");
            }
        }
    }
}