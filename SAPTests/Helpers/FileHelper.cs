namespace SAPTests.Helpers
{
    public static class FileHelper
    {
        /// <summary>
        /// Combines the base directory of the application with the specified relative path and returns the full path.
        /// </summary>
        /// <param name="relativePath">The relative path to combine with the base directory.</param>
        /// <returns>The full path based on the base directory and the relative path.</returns>
        public static string GetFullPathFromBase(string relativePath)
        {
            try
            {
                // Combine the base directory with the relative path and get the full path
                return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath));
            }
            catch (Exception ex)
            {
                // Log the exception (if you have a logging mechanism)
                // For example: Logger.LogError(ex, "Failed to get full path from base directory.");

                // Rethrow or handle the exception as needed
                throw new InvalidOperationException("Error while getting the full path from the base directory.", ex);
            }
        }
    }
}
