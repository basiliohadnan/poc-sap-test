using SAPTests.Helpers;
using SAPTests.MaxCompra.Administracao.Compras;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using Starline;
using static SAPTests.Helpers.ElementHandler;

namespace SAPTests.MaxCompra
{
    [TestClass]
    public class MaxCompraTests : WinAppDriver
    {
        protected string dataFilePath;
        protected ElementHandler elementHandler;
        protected string sheet;

        public MaxCompraTests()
        {
            sheet = "MaxComprasInit";
            dataFilePath = FileHelper.GetFullPathFromBase(Path.Combine("..", "..", "..", "..", "SAPTests", "dataset", "SAP.xlsx"));
            GetAppConfig();
            elementHandler = new ElementHandler();
        }

        private void GetAppConfig()
        {
            string queryName = "GetAppConfig";
            DataFetch dataFetch = new DataFetch(ConnType: "Excel", ConnXLS: dataFilePath);
            dataFetch.NewQuery(
                QueryName: queryName,
                    QueryText: $"SELECT * FROM [Config$]"
                    );

            Global.appPath = dataFetch.GetValue("APPPATH", queryName);
            Global.app = dataFetch.GetValue("APP", queryName);
        }

        protected void Initialize()
        {
            StartWinAppDriver();
            InitializeWinSession();
            InitializeAppSession(Global.appPath);
        }

        protected void Authenticate(string matricula, string filial = "000 - MATRIZ")
        {
            FillField(matricula);
            if (filial != "000 - MATRIZ")
            {
                WindowsElement lojasButton = FindElementByName("Open");
                lojasButton.Click();
                FillField(filial);
            }
            KeyPresser.PressKey("RETURN");
        }

        protected void SetAppSession()
        {
            string className = "Centura:MDIFrame";
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

        private void FillCredentials(DataFetch dataFetch, string queryName, int filialIndex = 0)
        {
            string stepDescription = "Realizar login do analista";
            string paramName = "matricula";
            string matricula = dataFetch.GetValue("MATRICULA", queryName);
            List<string> filiais = StringHandler.ParseStringToList(dataFetch.GetValue("FILIAIS", queryName));
            string paramValue = matricula;
            string expectedResult = "Login efetuado.";
            string printFileName;
            int lgsID;

            lgsID = Global.processTest.StartStep(stepDescription, logMsg: $"Tentando {stepDescription}",
                paramName: paramName, paramValue: paramValue);
            Authenticate(matricula, filiais[filialIndex]);
            SetAppSession();

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

        private void ValidateMainScreenShown()
        {
            string stepDescription = "Validar tela principal exibida";
            string paramName = "";
            string paramValue = "";
            string expectedResult = "Tela principal exibida";
            string printFileName;
            int lgsID = Global.processTest.StartStep(stepDescription, logMsg: $"Tentando {stepDescription}", paramName: paramName, paramValue: paramValue);

            try
            {
                string mainWindowClassName = "Centura:MDIFrame";
                WindowsElement mainWindow = FindElementByClassName(mainWindowClassName);
                Assert.IsNotNull(mainWindow);
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, printPath: printFileName, logMsg: expectedResult);
            }
            catch
            {
                printFileName = Global.processTest.CaptureWholeScreen();
                Global.processTest.EndStep(lgsID, status: "erro", printPath: printFileName, logMsg:
                    $"Erro ao tentar {stepDescription}.");
                throw new Exception($"Erro ao tentar {stepDescription}.");
            }
        }

        protected void Login(DataFetch dataFetch, string testName)
        {
            OpenApp();
            FillCredentials(dataFetch, testName);
            ValidateMainScreenShown();
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