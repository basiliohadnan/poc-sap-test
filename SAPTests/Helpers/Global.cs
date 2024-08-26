using OpenQA.Selenium.Appium.Windows;
using Starline;

namespace SAPTests.Helpers
{
    public class Global
    {
        public static string logonUser = Environment.UserName;
        public static WindowsDriver<WindowsElement> winSession;
        public static WindowsElement mainElement;
        public static WindowsDriver<WindowsElement> appSession;
        public static string app;
        public static string screenshotsDirectory;
        public static ProcessTest processTest = new ProcessTest();
        public static string customerName = "Customer";
        public static string suiteName = "suiteName";
        public static bool firstRun = true;
        public static DataFetch dataFetch;
        public static string appPath;
    }
}
