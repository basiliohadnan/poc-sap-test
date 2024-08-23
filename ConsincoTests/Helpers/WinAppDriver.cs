using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static SAPTests.Helpers.ElementHandler;
using Starline;
using SAPTests.Helpers;

namespace SAPTests.Helpers
{
    public class WinAppDriver
    {
        protected string dataSetFilePath = FileHelper.GetFullPathFromBase(Path.Combine("..", "..", "..", "..", "ConsincoTests", "dataset", "GerenciadordeCompras.xlsx"));

        protected void StartWinAppDriver()
        {
            string queryName = "GetAppConfig";
            DataFetch dataFetch = new DataFetch(ConnType: "Excel", ConnXLS: dataSetFilePath);
            dataFetch.NewQuery(
                QueryName: queryName,
                    QueryText: $"SELECT * FROM [Config$]"
                    );

            string winAppDriverPath = dataFetch.GetValue("WINAPPDRIVERPATH", queryName);

            if (Process.GetProcessesByName("WinAppDriver").Length == 0)
            {
                Process.Start(winAppDriverPath);
            }
        }

        protected void InitializeWinSession()
        {
            if (Global.mainElement == null)
            {
                AppiumOptions winCapabilities = new AppiumOptions();
                winCapabilities.AddAdditionalCapability("app", "Root");
                Global.winSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), winCapabilities);
                Global.mainElement = Global.winSession.FindElementByName("Desktop");
            }
        }

        protected void StopWinAppDriver()
        {
            foreach (Process process in Process.GetProcessesByName("WinAppDriver"))
            {
                process.Kill();
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            Global.appSession?.Quit();
            Global.winSession?.Quit();
            Global.mainElement = null;
            StopWinAppDriver();
        }

        protected void InitializeAppSession(string appPath)
        {
            Process process = Process.Start(appPath);
            string appName = "Conexão de Sistemas Consinco";
            WindowsElement appWindow = FindElementByName(appName, session: Global.winSession);
            Assert.IsNotNull(appWindow);

            // Get the window handle of the app's process
            nint mainWindowHandle = process.MainWindowHandle;
            
            // Identify the root level window of the app's process
            AppiumOptions rootCapabilities = new AppiumOptions();

            // Use the window handle as the appTopLevelWindow capability
            rootCapabilities.AddAdditionalCapability("appTopLevelWindow", mainWindowHandle.ToInt64().ToString("x"));
            Global.appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);
        }

        protected void SetAppSession(string className)
        {
            var appWindow = FindElementByClassName(className, session: Global.winSession);
            AppiumOptions appCapabilities = new AppiumOptions();
            var rootTopLevelWindowHandle = appWindow.GetAttribute("NativeWindowHandle");
            rootTopLevelWindowHandle = (int.Parse(rootTopLevelWindowHandle)).ToString("x"); // Convert to Hex
            appCapabilities.AddAdditionalCapability("appTopLevelWindow", rootTopLevelWindowHandle);
            Global.appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
        }

        public static void WaitSeconds(int seconds)
        {
            Thread.Sleep(seconds * 1000);
        }

        public static void FillField(string information)
        {
            Global.appSession.Keyboard.SendKeys(information);
        }
    }
}
