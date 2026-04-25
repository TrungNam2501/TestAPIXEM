using System;
using System.Linq;

namespace ServiceBBAPI.Helpers
{
    public static class BatchCalculator
    {
        private const double Epsilon = 0.000001;

        public static string Calculate(double kgStandard, double kgInput, double kgAlreadyPrinted, double limitKg)
        {
            int start = (int)Math.Floor(kgAlreadyPrinted / kgStandard) + 1;
            int end = (int)Math.Floor((kgAlreadyPrinted + kgInput - Epsilon) / kgStandard) + 1;

            int maxBatch = (int)Math.Ceiling(limitKg / kgStandard);
            if (end > maxBatch) end = maxBatch;

            if (start == end)
                return start.ToString();

            return string.Join("-", Enumerable.Range(start, end - start + 1));
        }
    }
}
