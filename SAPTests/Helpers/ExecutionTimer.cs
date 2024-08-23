using System.Diagnostics;

namespace SAPTests.Helpers
{
    public class ExecutionTimer
    {
        private Stopwatch stopwatch;

        public ExecutionTimer()
        {
            stopwatch = new Stopwatch();
        }

        public void Start()
        {
            stopwatch.Reset();
            stopwatch.Start();
        }

        public TimeSpan Stop()
        {
            stopwatch.Stop();
            return stopwatch.Elapsed;
        }
    }
}
