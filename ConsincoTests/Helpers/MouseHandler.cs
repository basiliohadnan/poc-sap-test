using SAPTests.Helpers;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;

namespace SAPTests.Helpers
{
    public class MouseHandler
    {
        public static void Click(BoundingRectangle boundingRectangle, int clickCount = 1)
        {
            int maxRetries = 3;
            int attempts = 0;
            while (attempts < maxRetries)
            {
                try
                {
                    for (int i = 0; i < clickCount; i++)
                    {
                        // Extract coordinates from the bounding rectangle
                        int offsetX = (boundingRectangle.Left + boundingRectangle.Right) / 2;
                        int offsetY = (boundingRectangle.Top + boundingRectangle.Bottom) / 2;

                        Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
                        Global.winSession.Mouse.Click(null);
                    }
                    return;
                }
                catch (WebDriverException e)
                {
                    if (e.Message.Contains("timed out"))
                    {
                        attempts++;
                        if (attempts == maxRetries)
                        {
                            throw;
                        }
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        public static void Click(int offsetX, int offsetY, int maxRetries = 5)
        {
            int attempts = 0;
            while (attempts < maxRetries)
            {
                try
                {
                    Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
                    Global.winSession.Mouse.Click(null);
                    return;
                }
                catch (WebDriverException e)
                {
                    if (e.Message.Contains("timed out"))
                    {
                        attempts++;
                        if (attempts == maxRetries)
                        {
                            throw;
                        }
                    }
                    else
                    {
                        throw;
                    }
                }
            }
        }

        public static void Click(WindowsElement element)

        {
            new Actions(Global.appSession).MoveToElement(element).Click().Perform();
        }

        public static void Click()

        {
            Global.appSession.Mouse.Click(null);
        }

        public static void DoubleClick(WindowsElement element)

        {
            new Actions(Global.appSession).MoveToElement(element).DoubleClick().Perform();
        }

        public static void DoubleClick()

        {
            Global.appSession.Mouse.DoubleClick(null);
        }

        public static void DoubleClickOn(BoundingRectangle boundingRectangle)
        {
            int offsetX = (boundingRectangle.Left + boundingRectangle.Right) / 2;
            int offsetY = (boundingRectangle.Top + boundingRectangle.Bottom) / 2;

            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
            Global.winSession.Mouse.DoubleClick(null);
        }

        public static void DoubleClick(int offsetX, int offsetY)
        {
            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);
            Global.winSession.Mouse.DoubleClick(null);
        }

        public static void DoubleTapOn(AppiumWebElement element)
        {
            TouchAction action = new TouchAction(Global.appSession);
            action.Tap(element).Perform();
            action.Tap(element).Perform();
        }

        public static void RightClick(WindowsElement element)
        {
            // Create an instance of the Actions class
            Actions actions = new Actions(Global.appSession);

            // Perform a right-click (context click) on the provided element
            actions.ContextClick(element).Perform();
        }

        public static void RightClick(BoundingRectangle boundingRectangle)
        {
            // Extract coordinates from the bounding rectangle
            int offsetX = (boundingRectangle.Left + boundingRectangle.Right) / 2;
            int offsetY = (boundingRectangle.Top + boundingRectangle.Bottom) / 2;

            // Move the mouse to the center of the bounding rectangle using the same reference point as in Click method
            Global.winSession.Mouse.MouseMove(Global.mainElement.Coordinates, offsetX, offsetY);

            // Create an instance of Actions class
            Actions actions = new Actions(Global.winSession);

            // Perform a right-click at the current position
            actions.ContextClick().Perform();
        }

        public static void RightClick(int x, int y)
        {
            // Create an instance of Actions class
            Actions actions = new Actions(Global.winSession);

            // Move the mouse to the specified coordinates and perform a right-click
            actions.MoveByOffset(x, y).ContextClick().Perform();
        }
    }
}
