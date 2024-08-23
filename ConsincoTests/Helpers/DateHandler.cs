namespace SAPTests.Helpers
{
    public class DateHandler
    {
        public static DateTime GetTodaysDate()
        {
            return DateTime.Today;
        }

        public static DateTime GetYesterdaysDate()
        {
            return DateTime.Today.AddDays(-1);
        }
    }
}
