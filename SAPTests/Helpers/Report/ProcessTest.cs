//================================================================================
// Class   : ProcessTest
// Version : 2.02
//
// Created : 20/04/2019 - 1.00 - Carlos Oliveira - Class creation
// Updated : 13/09/2019 - 2.00 - Gelder Carvalho - Class refactoring
// Updated : 04/05/2020 - 2.01 - Gelder Carvalho - New function Print for headless mode
// Updated : 08/09/2020 - 2.02 - Gelder Carvalho - Added Selenium helpers functions
//================================================================================

using System.Drawing;
using System.Drawing.Imaging;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using SAPTests.Helpers;

namespace Starline
{
    public class ProcessTest : AutoReport
    {
        public int defaultWidth = 1366;
        public int defaultHeight = 768;

        //Compatibility with older codes
        public string PathDoProjeto { get; set; }
        public string PastaBase { get; set; }

        public ProcessTest()
        { }

        public string PrintPageComSelenium(IWebDriver driver, bool full = false, int sleep = 0)
        {
            try
            {
                if (sleep > 0)
                {
                    System.Threading.Thread.Sleep(sleep * 1000);
                }

                //Create base directory and set image filename
                string printPath = GetAppPath() + "/Prints" + "/" + CustomerName + "/" + RptID.ToString() + "/" + SuiteName + "/" + ScenarioName + "/" + TestName;
                if (!Directory.Exists(printPath))
                {
                    Directory.CreateDirectory(printPath);
                }
                string filename = printPath + "/" + StepNumber.ToString().PadLeft(4, '0') + "_" + StepTurn.ToString().PadLeft(2, '0') + "-" + TestName.Replace("-", "_") + ".png";
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }

                if (full)
                {
                    /*
                        Dispositivo			Width	X	Heigh
                        Reponsive			400		X	1397
                        Galaxy S5			360		X	640
                        Pixel 2				411		X	731
                        Pixel 2 XL			411		X	823
                        iPhone 5/SE			320		X	568
                        iPhone 6/7/8		375		X	667
                        iPhone 6/7/8 Plus	414		X	736
                        iPhone X			375		X	812
                        iPad				768		X	1024
                        iPad Pro			1024	X	1366
                    */

                    Bitmap stitchedImage = null;

                    // First scroll to load all page components
                    ((IJavaScriptExecutor)driver).ExecuteScript(String.Format("window.scrollBy({0}, {1})", 0, -100000));
                    for (int p = 0; p < 200; p++)
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript(String.Format("window.scrollBy({0}, {1})", 0, 500));
                        System.Threading.Thread.Sleep(5);
                    }
                    ((IJavaScriptExecutor)driver).ExecuteScript(String.Format("window.scrollBy({0}, {1})", 0, -100000));
                    System.Threading.Thread.Sleep(200);

                    // Get full page size
                    long totalwidth1 = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.offsetWidth");//documentElement.scrollWidth");
                    long totalHeight1 = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.parentNode.scrollHeight");
                    int totalWidth = (int)totalwidth1;
                    int totalHeight = (int)totalHeight1;

                    // Get viewport size
                    long viewportWidth1 = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.clientWidth");//documentElement.scrollWidth");
                    long viewportHeight1 = (long)((IJavaScriptExecutor)driver).ExecuteScript("return window.innerHeight");//documentElement.scrollWidth");
                    int viewportWidth = (int)viewportWidth1;
                    int viewportHeight = (int)viewportHeight1;

                    // Split screen in multiple rectangles
                    List<Rectangle> rectangles = new List<Rectangle>();

                    // Loop until total height
                    for (int i = 0; i < totalHeight; i += viewportHeight)
                    {
                        int newHeight = viewportHeight;

                        // Fix if element height too big
                        if (i + viewportHeight > totalHeight)
                        {
                            newHeight = totalHeight - i;
                        }

                        // Loop until total width
                        for (int ii = 0; ii < totalWidth; ii += viewportWidth)
                        {
                            int newWidth = viewportWidth;

                            // Fix if element width too big
                            if (ii + viewportWidth > totalWidth)
                            {
                                newWidth = totalWidth - ii;
                            }

                            // Create and add new rectangle
                            Rectangle currRect = new Rectangle(ii, i, newWidth, newHeight);
                            rectangles.Add(currRect);
                        }
                    }

                    // Build image
                    stitchedImage = new Bitmap(totalWidth, totalHeight);

                    // Get all screenshots together
                    Rectangle previous = Rectangle.Empty;
                    foreach (var rectangle in rectangles)
                    {
                        // Calculate needed scrolling
                        if (previous != Rectangle.Empty)
                        {
                            int xDiff = rectangle.Right - previous.Right;
                            int yDiff = rectangle.Bottom - previous.Bottom;
                            ((IJavaScriptExecutor)driver).ExecuteScript(String.Format("window.scrollBy({0}, {1})", xDiff, yDiff));
                            System.Threading.Thread.Sleep(200);
                        }

                        // Take screenshot
                        var screenshot = ((ITakesScreenshot)driver).GetScreenshot();

                        // Build an image from the screenshot
                        Image screenshotImage;
                        using (MemoryStream memStream = new MemoryStream(screenshot.AsByteArray))
                        {
                            screenshotImage = Image.FromStream(memStream);
                        }

                        // Calculate source rectangle
                        Rectangle sourceRectangle = new Rectangle(viewportWidth - rectangle.Width, viewportHeight - rectangle.Height, rectangle.Width, rectangle.Height);

                        // Copy image
                        using (Graphics g = Graphics.FromImage(stitchedImage))
                        {
                            g.DrawImage(screenshotImage, rectangle, sourceRectangle, GraphicsUnit.Pixel);
                        }

                        // Set previous rectangle
                        previous = rectangle;
                    }

                    // Save image file
                    stitchedImage.Save(filename, ImageFormat.Png);
                }
                else
                {
                    Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                    ss.SaveAsFile(filename, ScreenshotImageFormat.Png);
                }
                return filename;
            }
            catch (Exception ex)
            {
                Print("Exception at PrintPageComSelenium", ex);
                return null;
            }
        }

        public string Print(IWebDriver driver, bool full = false, int sleep = 0, int customWidth = -1, int customHeight = 0)
        {
            try
            {
                if (sleep > 0)
                {
                    System.Threading.Thread.Sleep(sleep * 1000);
                }

                //Create base directory and set image filename
                string printPath = GetAppPath() + "/Prints" + "/" + CustomerName + "/" + RptID.ToString() + "/" + SuiteName + "/" + ScenarioName + "/" + TestName;
                if (!Directory.Exists(printPath))
                {
                    Directory.CreateDirectory(printPath);
                }
                string filename = printPath + "/" + StepNumber.ToString().PadLeft(4, '0') + "_" + StepTurn.ToString().PadLeft(2, '0') + "-" + TestName.Replace("-", "_") + ".png";
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }

                if (full)
                {
                    int originalWidth = driver.Manage().Window.Size.Width;
                    int originalHeight = driver.Manage().Window.Size.Height;
                    int totalWidth = 0;
                    int totalHeight = 0;

                    if (customWidth == -1)
                    {
                        totalWidth = defaultWidth;
                    }
                    else if (customWidth == 0)
                    {
                        long autoWidth = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.offsetWidth");//documentElement.scrollWidth");
                        totalWidth = (int)autoWidth;
                    }
                    else if (customWidth != 0)
                    {
                        totalWidth = customWidth;
                    }

                    if (customHeight == -1)
                    {
                        totalHeight = defaultHeight;
                    }
                    else if (customHeight == 0)
                    {
                        long autoHeight = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.parentNode.scrollHeight");
                        totalHeight = (int)autoHeight;
                    }
                    else if (customHeight != 0)
                    {
                        totalHeight = customHeight;
                    }

                    driver.Manage().Window.Size = new Size(totalWidth, totalHeight);
                    //Console.WriteLine(driver.Manage().Window.Size);

                    Screenshot ssFull = ((ITakesScreenshot)driver).GetScreenshot();
                    ssFull.SaveAsFile(filename, ScreenshotImageFormat.Png);

                    driver.Manage().Window.Size = new Size(originalWidth, originalHeight);
                    //Console.WriteLine(driver.Manage().Window.Size);
                }
                else
                {
                    Screenshot ssNormal = ((ITakesScreenshot)driver).GetScreenshot();
                    ssNormal.SaveAsFile(filename, ScreenshotImageFormat.Png);
                }

                return filename;
            }
            catch (Exception ex)
            {
                Print("Exception at Print", ex);
                return null;
            }
        }

        public string Screenshot(IWebDriver driver, string filename, bool full = false, int sleep = 0, int customWidth = -1, int customHeight = 0)
        {
            try
            {
                if (sleep > 0)
                {
                    System.Threading.Thread.Sleep(sleep);
                }

                //Create base directory and set image filename
                string printPath = GetAppPath() + "/Prints";
                if (!Directory.Exists(printPath))
                {
                    Directory.CreateDirectory(printPath);
                }
                string fullFilename = printPath + "/" + filename + ".png";
                if (File.Exists(fullFilename))
                {
                    File.Delete(fullFilename);
                }

                if (full)
                {
                    int originalWidth = driver.Manage().Window.Size.Width;
                    int originalHeight = driver.Manage().Window.Size.Height;
                    int totalWidth = 0;
                    int totalHeight = 0;

                    if (customWidth == -1)
                    {
                        totalWidth = defaultWidth;
                    }
                    else if (customWidth == 0)
                    {
                        long autoWidth = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.offsetWidth");//documentElement.scrollWidth");
                        totalWidth = (int)autoWidth;
                    }
                    else if (customWidth != 0)
                    {
                        totalWidth = customWidth;
                    }

                    if (customHeight == -1)
                    {
                        totalHeight = defaultHeight;
                    }
                    else if (customHeight == 0)
                    {
                        long autoHeight = (long)((IJavaScriptExecutor)driver).ExecuteScript("return document.body.parentNode.scrollHeight");
                        totalHeight = (int)autoHeight;
                    }
                    else if (customHeight != 0)
                    {
                        totalHeight = customHeight;
                    }

                    driver.Manage().Window.Size = new Size(totalWidth, totalHeight);
                    //Console.WriteLine(driver.Manage().Window.Size);

                    Screenshot ssFull = ((ITakesScreenshot)driver).GetScreenshot();
                    ssFull.SaveAsFile(fullFilename, ScreenshotImageFormat.Png);

                    driver.Manage().Window.Size = new Size(originalWidth, originalHeight);
                    //Console.WriteLine(driver.Manage().Window.Size);
                }
                else
                {
                    Screenshot ssNormal = ((ITakesScreenshot)driver).GetScreenshot();
                    ssNormal.SaveAsFile(fullFilename, ScreenshotImageFormat.Png);
                }

                return fullFilename;
            }
            catch (Exception ex)
            {
                Print("Exception at Print", ex);
                return null;
            }
        }

        public string CaptureWholeScreen(int sleep = 0)
        {
            try
            {
                if (sleep > 0)
                {
                    Thread.Sleep(sleep * 1000);
                }

                string printPath = GetAppPath() + "/Screenshots" + "/" + CustomerName + "/" + RptID.ToString() + "/" + SuiteName + "/" + ScenarioName + "/" + TestName;
                if (!Directory.Exists(printPath))
                {
                    Directory.CreateDirectory(printPath);
                }

                string fullFilename = printPath + "/" + GetStepNumber(StepName).ToString().PadLeft(4, '0') + "_" + StepTurn.ToString().PadLeft(2, '0') + "-" + TestName.Replace("-", "_") + ".jpg"; // Adjusted file extension to .jpg
                if (File.Exists(fullFilename))
                {
                    File.Delete(fullFilename);
                }

                // Capture screenshot using ITakesScreenshot interface
                var screenshot = ((ITakesScreenshot)Global.winSession).GetScreenshot();
                string pngFilePath = fullFilename; // Path to the PNG file
                screenshot.SaveAsFile(pngFilePath, ScreenshotImageFormat.Png);

                // Convert PNG to JPG
                string jpgFilePath = fullFilename.Replace(".png", ".jpg"); // Adjust file extension
                fullFilename = ImageEditor.ConvertPngToJpg(pngFilePath, jpgFilePath);

                return fullFilename;
            }
            catch (Exception ex)
            {
                Print("Exception at Print", ex);
                return null;
            }
        }

        public void Wait(int milliseconds)
        {
            Thread.Sleep(milliseconds);
        }

        public void Navigate(IWebDriver driver, string url)
        {
            driver.Navigate().GoToUrl(url);
            WaitForPageLoad(driver);
        }

        public void WaitForPageLoad(IWebDriver driver)
        {
            OpenQA.Selenium.Support.UI.WebDriverWait Wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30));
            try
            {
                Wait.Until(driver => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
            }
            catch (Exception)
            {

            }
        }

        public void WaitVisible(IWebDriver driver, By element)
        {
            OpenQA.Selenium.Support.UI.WebDriverWait wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(element));
            Wait(400);
        }

        public void WaitClickable(IWebDriver driver, By element)
        {
            OpenQA.Selenium.Support.UI.WebDriverWait wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementToBeClickable(element));
        }

        public void WaitSelectable(IWebDriver driver, By element)
        {
            OpenQA.Selenium.Support.UI.WebDriverWait wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementToBeSelected(element));
        }

        public void WaitExists(IWebDriver driver, By element)

        {
            OpenQA.Selenium.Support.UI.WebDriverWait wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementExists(element));
        }

        public void Click(IWebDriver driver, By element, int milliseconds = 0)
        {
            WaitClickable(driver, element);
            driver.FindElement(element).Click();
            Wait(milliseconds);
        }

        public void DoubleClick(IWebDriver driver, By element, int milliseconds = 0)
        {
            WaitClickable(driver, element);

            OpenQA.Selenium.Interactions.Actions actions = new OpenQA.Selenium.Interactions.Actions(driver);
            IWebElement webElement = driver.FindElement(element);
            actions.DoubleClick(webElement).Perform();
            Wait(milliseconds);
        }

        public void Clear(IWebDriver driver, By element)
        {
            WaitExists(driver, element);
            driver.FindElement(element).Clear();
        }

        public void SendKeys(IWebDriver driver, By element, string text, int milliseconds = 0)
        {
            WaitExists(driver, element);
            driver.FindElement(element).SendKeys(text);
            Wait(milliseconds);
        }

        public void SwitchFrame(IWebDriver driver, string frame)
        {
            driver.SwitchTo().ParentFrame();
            driver.SwitchTo().Frame(frame);
        }

        public void DefaultContentFrame(IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
        }

        public bool GetElementSelected(IWebDriver driver, By element)
        {
            WaitExists(driver, element);
            return driver.FindElement(element).Selected;
        }

        public string GetElementValue(IWebDriver driver, By element)
        {
            WaitExists(driver, element);
            return driver.FindElement(element).GetAttribute("value");
        }

        public string GetElementAttribute(IWebDriver driver, By element, string name)
        {
            WaitExists(driver, element);
            string result = "";
            IWebElement webElement = driver.FindElement(element);
            Size elementSize = webElement.Size;
            switch (name)
            {
                case "width":
                    result = elementSize.Width.ToString();
                    break;
                case "height":
                    result = elementSize.Height.ToString();
                    break;
                default:
                    result = driver.FindElement(element).GetAttribute(name);
                    break;
            }
            return result;
        }

        public string GetElementCssProperty(IWebDriver driver, By element, string property)
        {
            WaitExists(driver, element);
            return driver.FindElement(element).GetCssValue(property);
        }

        public string GetElementText(IWebDriver driver, By element)
        {
            WaitExists(driver, element);
            return driver.FindElement(element).Text;
        }

        public int GetElementsCount(IWebDriver driver, By element)
        {
            return driver.FindElements(element).Count;
        }

        public void SetElementAttribute(IWebDriver driver, By element, string name, string value)
        {
            WaitExists(driver, element);
            IWebElement webElement = driver.FindElement(element);
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0]." + name + "='" + value + "';", webElement);
        }

        public void SetSelectByValue(IWebDriver driver, By element, string value, int milliseconds = 0)
        {
            WaitExists(driver, element);
            var selectElement = new OpenQA.Selenium.Support.UI.SelectElement(driver.FindElement(element));
            selectElement.SelectByValue(value);
            Wait(milliseconds);
        }

        public void SetSelectByText(IWebDriver driver, By element, string text)
        {
            WaitExists(driver, element);
            var selectElement = new OpenQA.Selenium.Support.UI.SelectElement(driver.FindElement(element));
            selectElement.SelectByText(text);
        }

        public bool Exists(IWebDriver driver, By element)
        {
            bool result;
            try
            {
                IWebElement webElement = driver.FindElement(element);
                result = true;
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        public bool IsDisplayed(IWebDriver driver, By element)
        {
            bool result;
            try
            {
                IWebElement webElement = driver.FindElement(element);
                result = webElement.Displayed;
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        public void SendKeys(IWebDriver driver, string text, int milliseconds = 0)
        {
            OpenQA.Selenium.Interactions.Actions actions = new OpenQA.Selenium.Interactions.Actions(driver);
            actions.SendKeys(text).Build().Perform();
            Wait(milliseconds);
        }

        public void MouseOver(IWebDriver driver, By element, int milliseconds = 0)
        {
            IWebElement webElement = driver.FindElement(element);
            OpenQA.Selenium.Interactions.Actions actions = new OpenQA.Selenium.Interactions.Actions(driver);
            actions.MoveToElement(webElement).Perform();
            Wait(milliseconds);
        }

    }
}