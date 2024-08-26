using OpenQA.Selenium;
using SAPTests.Helpers;

namespace Starline
{
    public class ProcessTest : AutoReport
    {
        public int defaultWidth = 1366;
        public int defaultHeight = 768;

        public ProcessTest()
        { }

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
    }
}