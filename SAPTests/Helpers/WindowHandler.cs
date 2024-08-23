using SAPTests.Helpers;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using System.Collections.ObjectModel;
using OpenQA.Selenium;

namespace SAPTests.Helpers
{
    public class WindowHandler
    {
        public static void CloseWindow()
        {
            WindowsElement exitButton = ElementHandler.FindElementByName("Close", 3000);
            exitButton.Click();
        }

        public static void MaximizeWindow()
        {
            AppiumWebElement button = ElementHandler.FindElementByName("Maximize");
            button.Click();
        }

        public static void RestoreWindow()
        {
            ReadOnlyCollection<WindowsElement> buttons = ElementHandler.FindElementsByName("Restore");
            buttons[1].Click();
        }
    }
}
