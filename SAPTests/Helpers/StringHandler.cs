using System.Text.RegularExpressions;


namespace SAPTests.Helpers
{
    public class StringHandler
    {
        public static List<string> ParseStringToList(string inputString)
        {
            string[] parts = inputString.Split(',');
            List<string> itemList = new List<string>();

            foreach (string part in parts)
            {
                itemList.Add(part.Trim());
            }

            return itemList;
        }

        public static string RemoveCommasAndTrailingZeros(string input)
        {
            string pattern = @",0*"; // Pattern to match a comma followed by any number of zeros
            string replacement = ""; // Replace with an empty string
            return Regex.Replace(input, pattern, replacement);
        }

        public static string ConvertCommasToDots(string input)
        {
            return input.Replace(',', '.');
        }

        public static string CleanInput(string input)
        {
            // Remove all characters except digits and dots
            string cleaned = Regex.Replace(input, "[^0-9.]", "");

            // Find the position of the first dot
            int firstDotIndex = cleaned.IndexOf('.');

            if (firstDotIndex != -1)
            {
                // Keep only the first dot and remove any subsequent dots
                cleaned = cleaned.Substring(0, firstDotIndex + 1) + cleaned.Substring(firstDotIndex + 1).Replace(".", "");
            }

            return cleaned;
        }

        public static string ConvertListToString(List<string> list, string delimiter = ",")
        {
            if (list == null || list.Count == 0)
            {
                return string.Empty;
            }
            return string.Join(delimiter, list);
        }
    }
}
