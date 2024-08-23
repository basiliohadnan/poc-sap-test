using System.Runtime.InteropServices;

namespace SAPTests.Helpers
{
    public class ScreenHandler
    {
        [DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        public static Tuple<int, int> GetScreenResolution()
        {
            int screenWidth = GetSystemMetrics(0);  // SM_CXSCREEN
            int screenHeight = GetSystemMetrics(1); // SM_CYSCREEN

            return Tuple.Create(screenWidth, screenHeight);
        }
    }
}
