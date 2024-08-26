using SAPTests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Starline;
using static SAPTests.Helpers.ElementHandler;

namespace SAPTests.LoginTests
{
    [TestClass]
    public class LoginTests : WinAppDriver
    {
        protected string dataFilePath;
        protected ElementHandler elementHandler;
        protected string sheet;

        public LoginTests()
        {
            sheet = "login";
            dataFilePath = FileHelper.GetFullPathFromBase(Path.Combine("..", "..", "..", "..", "SAPTests", "dataset", "SAP.xlsx"));
            GetAppConfig();
            elementHandler = new ElementHandler();
        }

        private void GetAppConfig()
        {
            string testName = "GetAppConfig";
            DataFetch dataFetch = new DataFetch(ConnType: "Excel", ConnXLS: dataFilePath);
            dataFetch.NewQuery(
                QueryName: testName,
                    QueryText: $"SELECT * FROM [config$]"
                    );

            Global.appPath = dataFetch.GetValue("APPPATH", testName);
            Global.app = dataFetch.GetValue("APP", testName);
        }

        protected void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(Global.appPath);
        }

        protected void Authenticate(string username, string password)
        {
            FillField(username);
            KeyPresser.PressKey("TAB");
            FillField(password);
            KeyPresser.PressKey("RETURN");
        }

        protected void SetAppSession(string className)
        {
            SetAppSession(className);
        }

        protected void OpenMenu(string menuName)
        {
            WindowsElement menuItem = FindElementByName(menuName);
            menuItem.Click();
        }

        private void OpenApp()
        {
            string stepDescription = "Abrir app";
            int lgsID;
            string printFileName;
            string paramName = "appPath";
            string paramValue = Global.appPath;
            string expectedResult = "App aberto.";

            lgsID = Global.processTest.StartStep(stepDescription, logMsg:
                $"Tentando {stepDescription}", paramName: paramName, paramValue: paramValue);
            Initialize();
            printFileName = Global.processTest.CaptureWholeScreen();
            string welcomeWindowName = "Conexão de Sistemas SAP";
            WindowsElement welcomeWindow = FindElementByName(welcomeWindowName);
            Assert.IsNotNull(welcomeWindow);
            if (welcomeWindow == null)
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar {stepDescription}.");
                throw new Exception($"Erro ao tentar {stepDescription}.");
            }
            Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: expectedResult);
        }

        private void FillCredentials(DataFetch dataFetch, string testName, int filialIndex = 0)
        {
            string stepDescription = "Realizar login do usuário";
            string username = dataFetch.GetValue("USERNAME", testName);
            string password = dataFetch.GetValue("PASSWORD", testName);
            string className = dataFetch.GetValue("CLASSNAME", testName);
            string expectedResult = "Login efetuado.";
            string printFileName;
            int lgsID;

            lgsID = Global.processTest.StartStep(stepDescription, logMsg: $"Tentando {stepDescription}",
                paramName: "username, password", paramValue: $"{username}, {password}");
            Authenticate(username, password);
            SetAppSession(className);

            string appName = "Application";
            WindowsElement appMenu = FindElementByName(appName);
            Assert.IsNotNull(appMenu);

            if (appMenu == null)
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar {stepDescription}.");
                throw new Exception($"Erro ao tentar {stepDescription}.");

            }

            printFileName = Global.processTest.CaptureWholeScreen();
            Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: expectedResult);

            string databaseWarningName = "Não foi definido a versão do módulo no BANCO DE DADOS, para o sistema de Segurança";
            WindowsElement warning = FindElementByXPathPartialName(databaseWarningName);

            if (warning != null)
            {
                warning.Click();
                KeyPresser.PressKey("RETURN");

            }
        }

        protected void Login(DataFetch dataFetch, string testName)
        {
            OpenApp();
            FillCredentials(dataFetch, testName);
        }

        [TestMethod, TestCategory("done")]
        public void AbrirAplicacao()
        {
            string testName = TestHandler.GetCurrentMethodName();
            if (TestHandler.SetExecutionControl(dataFilePath, sheet, testName) == "")
            {
                TestHandler.StartTest(Global.dataFetch, testName);
                try
                {
                    TestHandler.DoTest(Global.dataFetch, testName);
                    TestHandler.DefineSteps(testName);
                    Login(Global.dataFetch, testName);
                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testName, "passed");
                }
                catch (Exception ex)
                {
                    ExcelHelper.UpdateTestResult(dataFilePath, sheet, testName, "failed");
                    throw;
                }
                finally
                {
                    TestHandler.EndTest(Global.dataFetch, testName);
                }
            }
        }
    }
}