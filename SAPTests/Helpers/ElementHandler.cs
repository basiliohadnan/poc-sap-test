using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;


namespace SAPTests.Helpers
{
    public class ElementHandler
    {
        public static ReadOnlyCollection<WindowsElement> FindElementsByClassName(string className)
        {
            const int maxAttempts = 10;
            int attempts = 1;

            while (attempts <= maxAttempts)
            {
                try
                {
                    // Find elements by class name
                    var elements = Global.appSession.FindElementsByClassName(className);

                    // Check if any elements are found
                    if (elements != null && elements.Count > 0)
                    {
                        // Return the list of elements found
                        return elements;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding elements by class name: {ex.Message}");
                    attempts++;
                }
            }

            // Elements not found after max attempts
            Console.WriteLine($"No elements with class name '{className}' found after {maxAttempts} attempts.");
            return null;
        }

        public static ReadOnlyCollection<WindowsElement> FindElementsByName(string name)
        {
            const int maxAttempts = 10;
            int attempts = 1;

            while (attempts <= maxAttempts)
            {
                try
                {
                    // Find elements by name
                    var elements = Global.appSession.FindElementsByName(name);

                    // Check if any elements are found
                    if (elements != null && elements.Count > 0)
                    {
                        // Return the list of elements found
                        return elements;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding elements by name: {ex.Message}");
                    attempts++;
                }
            }

            // Elements not found after max attempts
            Console.WriteLine($"No elements with name '{name}' found after {maxAttempts} attempts.");
            return null;
        }

        public static WindowsElement FindElementByName(string name, int milliseconds = 1000, WindowsDriver<WindowsElement> session = null)
        {
            if (session == null)
            {
                session = Global.appSession;
            }

            const int maxAttempts = 10;
            int attempts = 1;

            while (attempts <= maxAttempts)
            {
                Thread.Sleep(milliseconds);
                try
                {
                    WindowsElement element = session.FindElementByName(name);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception e)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by name {name}, attempt {attempts}: {e.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with name '{name}' not found after {maxAttempts} attempts.");
            return null;
        }

        public static WindowsElement FindElementByXPath(string xpath, int milliseconds = 1000, WindowsDriver<WindowsElement> session = null)
        {
            if (session == null)
            {
                session = Global.appSession;
            }

            const int maxAttempts = 10;
            int attempts = 1;

            while (attempts <= maxAttempts)
            {
                Thread.Sleep(milliseconds);
                try
                {
                    WindowsElement element = session.FindElementByXPath(xpath);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception e)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by xpath {xpath}, attempt {attempts}: {e.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with xpath '{xpath}' not found after {maxAttempts} attempts.");
            return null;
        }

        public static WindowsElement FindElementByClassName(string className, int milliseconds = 1000, WindowsDriver<WindowsElement> session = null)
        {
            if (session == null)
            {
                session = Global.appSession;
            }

            const int maxAttempts = 10;
            int attempts = 1;
            WindowsElement element = null;

            while (attempts <= maxAttempts)
            {
                try
                {
                    element = session.FindElementByClassName(className);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by class name: {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with class name '{className}' not found after {maxAttempts} attempts.");
            return null;
        }

        public static WindowsElement FindElementByAccessibilityId(string accessibilityId, int milliseconds = 1000, WindowsDriver<WindowsElement> session = null)
        {
            if (session == null)
            {
                session = Global.appSession;
            }

            const int maxAttempts = 10;
            int attempts = 1;
            WindowsElement element = null;

            while (attempts <= maxAttempts)
            {
                try
                {
                    element = session.FindElementByAccessibilityId(accessibilityId);
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element by accessibility ID: {ex.Message}");
                    attempts++;
                }

                // Wait for a short period before retrying
                System.Threading.Thread.Sleep(milliseconds / maxAttempts);
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with accessibility ID '{accessibilityId}' not found after {maxAttempts} attempts.");
            return null;
        }

        public static WindowsElement FindElementByXPathPartialName(string partialName)
        {
            const int maxAttempts = 10;
            int attempts = 1;
            WindowsElement element = null;

            while (attempts <= maxAttempts)
            {
                try
                {
                    element = Global.appSession.FindElement(By.XPath($"//*[contains(@Name, '{partialName}')]"));
                    if (element != null)
                    {
                        // Element found, return it
                        return element;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Element not found, continue trying
                    attempts++;
                }
                catch (Exception ex)
                {
                    // Log any other exceptions and retry
                    Console.WriteLine($"Exception occurred while finding element {partialName} by XPath : {ex.Message}");
                    attempts++;
                }
            }

            // Element not found after max attempts
            Console.WriteLine($"Element with partial name '{partialName}' not found after {maxAttempts} attempts.");
            return null;
        }

        public static bool VerifyCheckBoxIsOn(WindowsElement checkbox)
        {
            return checkbox.Selected;
        }

        public static void ConfirmWindow(string windowName, int buttonIndex = 0, int timeout = 1000)
        {
            WindowsElement foundWindow = FindElementByXPath(windowName, timeout);
            ReadOnlyCollection<AppiumWebElement> buttons = foundWindow.FindElementsByClassName("Button");
            AppiumWebElement button = buttons[buttonIndex];
            button.Click();
        }

        public static string ExtractTextWithPattern(string text, string pattern)
        {
            Match match = Regex.Match(text, pattern, RegexOptions.Singleline);
            if (match.Success)
            {
                // Extract the matched text after the colon
                string matchedText = match.Groups[1].Value.Trim();

                // Split the matched text by newline or space and filter out empty entries
                string[] parts = matchedText.Split(new[] { '\n', '\r', ' ' }, StringSplitOptions.RemoveEmptyEntries);

                // Join the parts with a comma
                return string.Join(",", parts);
            }
            return null;
        }

        public static string oldExtractTextWithPattern(string text, string pattern)
        {
            Match match = Regex.Match(text, pattern);
            if (match.Success)
            {
                return match.Value;
            }
            return null;
        }

        public static WindowsElement FindElementByClassAndName(string className, string name)
        {
            ReadOnlyCollection<WindowsElement> elements = FindElementsByClassName(className);
            WindowsElement element = elements.FirstOrDefault(b => b.GetAttribute("Name") == name);
            return element;
        }

        public static bool CanScrollDown(WindowsElement scrollbarElement)
        {
            try
            {
                // Get the current value, maximum, and minimum values of the scrollbar
                double currentValue = Convert.ToDouble(scrollbarElement.GetAttribute("RangeValue.Value"));
                double maximumValue = Convert.ToDouble(scrollbarElement.GetAttribute("RangeValue.Maximum"));
                double minimumValue = Convert.ToDouble(scrollbarElement.GetAttribute("RangeValue.Minimum"));

                // Determine if we can scroll down
                return currentValue < maximumValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred while checking scroll position: {ex.Message}");
                return false;
            }
        }
    }
}
