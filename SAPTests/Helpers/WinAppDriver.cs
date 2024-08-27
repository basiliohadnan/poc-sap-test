using System.Runtime.InteropServices;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static SAPTests.Helpers.ElementHandler;
using Starline;
using OpenQA.Selenium;
using System.Text;


namespace SAPTests.Helpers
{
    public class WinAppDriver
    {
        protected string dataFilePath = FileHelper.GetFullPathFromBase(Path.Combine("..", "..", "..", "..", "SAPTests", "dataset", "SAP.xlsx"));

        protected void StartWinAppDriver()
        {
            string queryName = "GetAppConfig";
            Global.dataFetch = new DataFetch(ConnType: "Excel", ConnXLS: dataFilePath);
            Global.dataFetch.NewQuery(
                QueryName: queryName,
                    QueryText: $"SELECT * FROM [config$]"
                    );

            string winAppDriverPath = Global.dataFetch.GetValue("WINAPPDRIVERPATH", queryName);

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

            // Wait until the process is fully started and the splash screen has passed
            WaitSeconds(10);

            // Find the correct main window handle
            nint mainWindowHandle = IntPtr.Zero;
            int maxRetries = 30; // Maximum retries (30 * 1s = 30 seconds)
            while (maxRetries > 0 && mainWindowHandle == IntPtr.Zero)
            {
                mainWindowHandle = FindMainWindowHandle(process);
                if (mainWindowHandle != IntPtr.Zero)
                    break;

                WaitSeconds(1);
                maxRetries--;
            }

            if (mainWindowHandle == IntPtr.Zero)
            {
                throw new Exception("Main window handle not found.");
            }

            // Identify the root level window of the app's process
            AppiumOptions rootCapabilities = new AppiumOptions();
            rootCapabilities.AddAdditionalCapability("appTopLevelWindow", mainWindowHandle.ToInt64().ToString("x"));
            Global.appSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), rootCapabilities);

            var everyElementFromAppSession = Global.appSession.FindElements(By.XPath("//*"));
        }

        // Method to find the main window handle of the process
        private nint FindMainWindowHandle(Process process)
        {
            nint windowHandle = IntPtr.Zero;
            foreach (ProcessThread thread in process.Threads)
            {
                EnumThreadWindows((uint)thread.Id, (hWnd, lParam) =>
                {
                    if (IsMainWindow(hWnd))
                    {
                        windowHandle = hWnd;
                        return false; // Stop enumeration
                    }
                    return true; // Continue enumeration
                }, IntPtr.Zero);

                if (windowHandle != IntPtr.Zero)
                    break;
            }
            return windowHandle;
        }

        // Helper method to determine if a window handle is the main application window
        private bool IsMainWindow(nint handle)
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                string windowTitle = Buff.ToString();
                // Ensure the window title corresponds to the SAP Logon main window
                return windowTitle.Contains("SAP Logon") && !windowTitle.Contains("Splash");
            }
            return false;
        }

        // P/Invoke declarations
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnumThreadWindows(uint dwThreadId, EnumThreadWndProc lpfn, nint lParam);

        delegate bool EnumThreadWndProc(nint hWnd, nint lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool IsWindowVisible(nint hWnd);


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
